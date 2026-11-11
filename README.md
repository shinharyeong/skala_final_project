# skala_final_project

SKALA 최종 프로젝트 모노레포입니다.

## 폴더 구조

| 폴더 | 내용 | 기술 |
| --- | --- | --- |
| [`frontend/`](frontend/) | 웹 프론트엔드 | Vue 3, Vite, Vue Router, Pinia |
| [`backend/`](backend/) | API 서버 | Spring Boot 3.4, Java 21, Gradle |
| [`ai/`](ai/) | AI 기능 | FastAPI, LangChain (구성 예정) |
| [`parser/`](parser/) | 경영기획 Excel 파싱·검증 | Python 3.11+ (표준 라이브러리) |
| [`docs/`](docs/) | 공통 문서 | |

## 로컬 실행

각 영역은 독립적으로 실행합니다. 자세한 내용은 폴더별 README를 참고하세요.

```bash
# frontend — Node.js 20.19+
cd frontend && npm install && npm run dev        # http://localhost:5173

# backend — JDK 21
cd backend && ./gradlew bootRun                  # http://localhost:8080

# parser — Python 3.11+
cd parser && PYTHONPATH=src python3 -B -m unittest discover -s tests -v
```

## 협업 규칙

브랜치 전략, 커밋·이슈·PR 규칙은 [CONTRIBUTING.md](CONTRIBUTING.md)를 따릅니다.

입력 데이터, 실행 결과, `.env` 등 비밀 정보는 커밋하지 않습니다.
