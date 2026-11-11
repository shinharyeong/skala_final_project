"""OpenAI Responses via stdlib: strict JSON, retries, cache, usage accounting.

Endpoint is fixed; spreadsheet content cannot choose a destination or tool.
The API key and the Authorization header we send are never logged. On an
HTTP error, only a sanitized diagnostic is surfaced: HTTP status, the
server's error.code/error.type/error.message, and the request id — never
the raw response body verbatim and never anything from the request we sent.
"""
import hashlib
import http.client
import json
import os
import random
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path

PROMPT_VERSION = "2026-11-11.3"

# Transient HTTP errors (429/5xx/408) get up to this many total attempts,
# guided by the server's own Retry-After. A connection failure or a client
# read timeout gets fewer: retrying an already-slow request blindly can
# just repeat the same multi-minute wait (and, if the server had already
# started billable processing, repeat that cost) for no new information.
HTTP_RETRY_MAX_ATTEMPTS = 4
CONNECTION_RETRY_MAX_ATTEMPTS = 2


class ModelError(RuntimeError):
    pass


def _error_detail(http_error, model):
    """Extract a safe, structured diagnostic from an OpenAI HTTPError.

    Reads only the server's response status/headers/JSON error body — never
    the request we sent, so the Authorization header and API key can never
    appear here. Missing/unparseable fields degrade to None rather than
    raising, so a malformed error body still yields a usable diagnostic.
    """
    request_id = None
    for name in ("x-request-id", "openai-request-id", "x-request-id-openai"):
        request_id = http_error.headers.get(name) if http_error.headers else None
        if request_id:
            break
    info = {}
    try:
        raw = http_error.read()
        parsed = json.loads(raw.decode("utf-8", errors="replace")) if raw else {}
        if isinstance(parsed, dict):
            info = parsed.get("error") or {}
    except Exception:
        info = {}
    return {
        "http_status": http_error.code,
        "request_id": request_id,
        "error_code": info.get("code") if isinstance(info, dict) else None,
        "error_type": info.get("type") if isinstance(info, dict) else None,
        "error_message": info.get("message") if isinstance(info, dict) else None,
        "model_requested": model,
    }


def read_env(path):
    """Load only documented variables, respecting explicit process environment."""
    if not path.is_file():
        return
    allowed = {"OPENAI_API_KEY", "PARSER_AI_MODEL", "PARSER_AI_REASONING_EFFORT", "PARSER_AI_MAX_OUTPUT_TOKENS", "PARSER_AI_TIMEOUT_SECONDS", "PARSER_AI_CONCURRENCY", "PARSER_AI_REVIEW"}
    for line in path.read_text().splitlines():
        key, sep, value = line.strip().partition("=")
        if sep and key in allowed:
            os.environ.setdefault(key, value.strip().strip("\"'"))


class ResponsesClient:
    def __init__(self, cache_dir, *, model=None, effort=None, timeout=None, max_output_tokens=None, transport=None, max_calls=300, progress=None):
        self.model = model or os.getenv("PARSER_AI_MODEL", "gpt-6-astra")
        self.effort = effort or os.getenv("PARSER_AI_REASONING_EFFORT", "high")
        self.timeout = timeout or float(os.getenv("PARSER_AI_TIMEOUT_SECONDS", "240"))
        self.max_output_tokens = max_output_tokens or int(os.getenv("PARSER_AI_MAX_OUTPUT_TOKENS", "32768"))
        self.cache_dir = Path(cache_dir)
        self.transport = transport or self._request
        self.max_calls = max_calls
        self.events = []
        self.lock = threading.Lock()
        self.call_count = 0
        self.http_post_attempts = 0
        # One line per request-start/attempt/completion/retry, so a slow
        # request is visible instead of looking hung. Never receives the key,
        # the Authorization header, or any workbook content — only task
        # names, attempt counts, elapsed seconds and sanitized error detail.
        self.progress = progress or (lambda message: None)
        self._context = threading.local()

    def _notify(self, message):
        try:
            self.progress(message)
        except Exception:
            pass  # progress reporting must never break the actual request

    def _task_label(self):
        return getattr(self._context, "task", "request")

    def _request(self, body):
        key = os.getenv("OPENAI_API_KEY")
        if not key:
            raise ModelError("OPENAI_API_KEY가 없습니다. parser_ai/.env에 설정한 뒤 run을 실행하세요. prepare는 키 없이 가능합니다.")
        request = urllib.request.Request(
            "https://api.openai.com/v1/responses", data=json.dumps(body).encode(),
            headers={"Authorization": "Bearer " + key, "Content-Type": "application/json"}, method="POST")
        task = self._task_label()
        http_attempt = 0
        connection_attempt = 0
        while True:
            start = time.monotonic()
            self._notify(f"[{task}] 요청 전송 (http 시도 {http_attempt + 1}/{HTTP_RETRY_MAX_ATTEMPTS}, "
                        f"연결 시도 {connection_attempt + 1}/{CONNECTION_RETRY_MAX_ATTEMPTS}, "
                        f"모델={self.model}, timeout={self.timeout:.0f}s)")
            try:
                with self.lock:
                    self.http_post_attempts += 1
                with urllib.request.urlopen(request, timeout=self.timeout) as response:
                    data = json.load(response)
                self._notify(f"[{task}] 응답 수신 ({time.monotonic() - start:.1f}s)")
                return data
            except urllib.error.HTTPError as error:
                elapsed = time.monotonic() - start
                detail = _error_detail(error, self.model)
                http_attempt += 1
                # Retry transient errors only. Never silently switch to a weaker model.
                if error.code not in {408, 429, 500, 502, 503, 504, 520} or http_attempt >= HTTP_RETRY_MAX_ATTEMPTS:
                    self._notify(f"[{task}] 실패 응답 ({elapsed:.1f}s, HTTP {error.code})")
                    raise ModelError(
                        "OpenAI API 오류(비밀값 제외): " + json.dumps(detail, ensure_ascii=False)
                        + ". 키·모델 접근 권한·프로젝트/조직 설정을 확인하세요. 모델은 자동으로 바꾸지 않습니다."
                    ) from None
                try:
                    delay = min(30, float(error.headers.get("Retry-After", 2 ** http_attempt)))
                except ValueError:
                    delay = 2 ** http_attempt
                self._notify(f"[{task}] 재시도 대기 {delay:.1f}s 예정 (HTTP {error.code}, {elapsed:.1f}s 경과)")
                time.sleep(max(0, delay) + random.random())
            except (urllib.error.URLError, TimeoutError, ConnectionError, http.client.IncompleteRead) as error:
                # NOTE: a connection dropped mid-response (e.g. RemoteDisconnected,
                # a ConnectionResetError subclass) is NOT wrapped into URLError by
                # urlopen() — getresponse() is called outside urlopen's own
                # try/except, and json.load(response) reads the body even later
                # still. Both raise the raw stdlib exception. This previously went
                # completely uncaught here and propagated as an opaque top-level
                # failure that discarded all prior progress for the file.
                elapsed = time.monotonic() - start
                connection_attempt += 1
                closed_mid_response = isinstance(error, (ConnectionError, http.client.IncompleteRead))
                kind = "연결이 응답 중간에 끊김" if closed_mid_response else "연결 실패 또는 응답 시간 초과"
                if connection_attempt >= CONNECTION_RETRY_MAX_ATTEMPTS:
                    self._notify(f"[{task}] {kind}로 중단 ({elapsed:.1f}s 경과, timeout={self.timeout:.0f}s, {type(error).__name__})")
                    raise ModelError(
                        f"OpenAI {kind}({elapsed:.1f}s, timeout={self.timeout:.0f}s, 모델={self.model}, "
                        f"오류 종류={type(error).__name__}). 같은 요청을 반복 재시도해 비용을 늘리지 않도록 연결 재시도 "
                        "횟수를 제한했습니다. 이미 진행 중이던 요청 자체를 강제로 취소할 수는 없습니다(표준 라이브러리 제약). "
                        "네트워크 연결, PARSER_AI_TIMEOUT_SECONDS, 입력 크기를 확인하세요."
                    ) from None
                self._notify(f"[{task}] {kind}로 재시도 예정 ({elapsed:.1f}s 경과, {type(error).__name__})")
                time.sleep(2 ** connection_attempt + random.random())

    def ask(self, task, instructions, context, schema, *, refresh=False):
        body = {
            "model": self.model, "store": False,
            "reasoning": {"effort": self.effort},
            "max_output_tokens": self.max_output_tokens,
            "input": [{"role": "system", "content": instructions},
                      {"role": "user", "content": json.dumps(context, ensure_ascii=False, separators=(",", ":"))}],
            "text": {"format": {"type": "json_schema", "name": task, "schema": schema, "strict": True}},
        }
        encoded = json.dumps(body, sort_keys=True, ensure_ascii=False)
        digest = hashlib.sha256(json.dumps({"version": PROMPT_VERSION, "body": body}, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
        path = self.cache_dir / (digest + ".json")
        if path.is_file() and not refresh:
            cached = json.loads(path.read_text())
            with self.lock:
                self.events.append({"task": task, "cached": True, "model": self.model, "seconds": 0, "usage": {}, "request_hash": digest})
            self._notify(f"[{task}] 캐시 사용 — API 호출 없음")
            return cached["data"]
        with self.lock:
            if self.call_count >= self.max_calls:
                raise ModelError(f"호출 상한 {self.max_calls}회에 도달했습니다. 캐시를 보존했으며 --max-calls로 조정할 수 있습니다.")
            self.call_count += 1
            call_number = self.call_count
        self._context.task = task
        self._notify(f"[{task}] 처리 시작 (누적 호출 {call_number}/{self.max_calls}, 요청 크기 {len(encoded) / 1024:.1f}KB)")
        start = time.monotonic()
        try:
            response = self.transport(body)
        except Exception as error:
            # Record what IS knowable about a failed attempt (no usage data
            # exists for a call that never returned a parseable response) so
            # "확인 가능한 API 사용량" still reflects failed calls, not only
            # successful ones.
            with self.lock:
                self.events.append({"task": task, "cached": False, "failed": True, "model": self.model,
                                    "error_type": type(error).__name__, "seconds": round(time.monotonic() - start, 3),
                                    "usage": {}, "request_hash": digest})
            raise
        self._notify(f"[{task}] 응답 수신, 검증 중 ({time.monotonic() - start:.1f}s)")
        if response.get("status") not in {None, "completed"}:
            raise ModelError("모델 응답이 완료되지 않았습니다. 출력 토큰 한도와 응답 상태를 확인하세요.")
        pieces = []
        for entry in response.get("output", []):
            for content in entry.get("content", []):
                if content.get("type") == "refusal":
                    raise ModelError("모델이 해당 요청에 대한 구조화 응답을 반환하지 않았습니다.")
                if content.get("type") == "output_text":
                    pieces.append(content["text"])
        try:
            data = json.loads("".join(pieces))
        except (ValueError, TypeError):
            raise ModelError("모델 JSON 응답을 읽을 수 없습니다. 원본 추출은 보존되며 룰 결과로 대체하지 않습니다.") from None
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        envelope = {"data": data, "model": response.get("model", self.model), "prompt_version": PROMPT_VERSION,
                    "usage": response.get("usage", {}), "request_hash": digest}
        temporary = path.with_suffix(f".{threading.get_ident()}.tmp")
        temporary.write_text(json.dumps(envelope, ensure_ascii=False), encoding="utf-8")
        temporary.replace(path)
        with self.lock:
            self.events.append({"task": task, "cached": False, "model": envelope["model"], "seconds": round(time.monotonic() - start, 3),
                                "usage": envelope["usage"], "request_hash": digest})
        self._notify(f"[{task}] 완료 (총 {time.monotonic() - start:.1f}s, 캐시에 저장)")
        return data
