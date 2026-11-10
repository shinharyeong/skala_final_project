"""프로젝트 경로를 검증하고 입력 파일 목록을 조회한다.

설정 파일과 경로 정보만 읽으며 원본 내용이나 디렉토리를 변경하지 않는다.
함수별 사용 관계는 docs/development/function-reference.md에서 관리한다.
"""

import tomllib
import unicodedata
from dataclasses import dataclass
from pathlib import Path

CONFIG_FILE = Path("config/paths.toml")
WORKBOOK_EXTENSIONS = frozenset({".xlsx", ".xlsb"})


def _paths_overlap(first: Path, second: Path) -> bool:
    """두 경로가 같거나 서로를 포함하는지 문자열 기준으로 비교한다.

    Args:
        first: 비교할 첫 번째 경로. 호출자가 해석을 마친 경로를 전달한다.
        second: 비교할 두 번째 경로. 첫 번째 경로와 같은 기준을 사용한다.

    Returns:
        대소문자와 유니코드 조합 차이를 무시했을 때 겹치면 True.

    Raises:
        없음. Path 타입의 정상 입력을 기준으로 한다.

    Side Effects:
        없음. 파일시스템에 접근하지 않으며 심볼릭 링크도 직접 해석하지 않는다.
    """
    # 운영체제가 달라도 같은 설정을 쓰도록 대소문자·유니코드 차이를 제거한다.
    left = tuple(unicodedata.normalize("NFC", part).casefold() for part in first.parts)
    right = tuple(
        unicodedata.normalize("NFC", part).casefold() for part in second.parts
    )
    common_length = min(len(left), len(right))
    return left[:common_length] == right[:common_length]


def _find_project_root() -> Path:
    """설정 파일과 패키지 정의가 함께 있는 프로젝트 루트를 찾는다.

    Args:
        없음. 현재 작업 위치와 이 소스 파일의 위치를 탐색 기준으로 사용한다.

    Returns:
        pyproject.toml과 config/paths.toml을 포함하는 루트의 절대 경로.

    Raises:
        FileNotFoundError: 모든 후보에서 프로젝트 루트를 찾지 못한 경우.
        OSError: 현재 위치 확인이나 경로 해석에서 파일시스템 오류가 전파된 경우.
        RuntimeError: 실행 환경에 따라 링크 순환 등 경로 해석이 실패한 경우.

    Side Effects:
        후보 경로의 파일 존재 여부를 확인한다. 내용 읽기와 파일 쓰기는 없다.
    """
    # 현재 위치의 프로젝트를 우선하고, 없을 때 이 소스가 속한 프로젝트를 찾는다.
    for start in (Path.cwd(), Path(__file__).resolve().parent):
        for candidate in (start, *start.parents):
            if (candidate / "pyproject.toml").is_file() and (
                candidate / CONFIG_FILE
            ).is_file():
                return candidate.resolve()
    raise FileNotFoundError(
        "프로젝트 루트를 찾을 수 없습니다. --root로 프로젝트 경로를 지정하세요."
    )


@dataclass(frozen=True)
class ProjectPaths:
    """입력·출력·로그 경로를 보관하는 불변 설정 객체.

    Attributes:
        root: 프로젝트 루트 경로.
        input_dir: 읽기 전용 입력 파일의 디렉토리 경로.
        output_dir: 향후 승인된 결과 저장에 사용할 경로.
        logs_dir: 향후 승인된 로그 저장에 사용할 경로.

    load()를 사용하면 검증된 절대 경로를 얻는다. 생성자를 직접 호출하면
    검증을 우회하므로 일반 호출자는 load()를 사용한다. 경로의 존재와
    저장 권한은 보장하지 않으며, 이 객체가 파일이나 폴더를 생성하지 않는다.
    """

    root: Path
    input_dir: Path
    output_dir: Path
    logs_dir: Path

    @classmethod
    def load(cls, root: str | Path | None = None) -> "ProjectPaths":
        """설정 파일을 읽어 서로 겹치지 않는 프로젝트 내부 경로를 반환한다.

        Args:
            root: 프로젝트 루트. None이면 자동 탐색한다. 상대 경로를 주면
                현재 작업 위치를 기준으로 해석하며 '~'는 사용자 홈으로 확장한다.

        Returns:
            root, input_dir, output_dir, logs_dir를 담은 ProjectPaths 객체.
            설정된 폴더의 실제 존재 여부나 접근 권한은 보장하지 않는다.

        Raises:
            FileNotFoundError: 프로젝트 루트나 경로 설정 파일을 찾지 못한 경우.
            tomllib.TOMLDecodeError: 설정 파일이 올바른 TOML 형식이 아닌 경우.
            ValueError: 필수 설정이 없거나 경로가 프로젝트 밖 또는 서로 겹친 경우.
            OSError: 설정 읽기나 경로 해석에서 파일시스템 오류가 전파된 경우.
            RuntimeError: 홈 경로나 심볼릭 링크를 실행 환경에서 해석할 수 없는 경우.

        Side Effects:
            config/paths.toml을 읽고 경로 정보를 조회한다.
            원본 파일 내용을 읽거나 파일·디렉토리를 생성하지 않는다.
        """
        project_root = (
            Path(root).expanduser().resolve()
            if root is not None
            else _find_project_root()
        )
        with (project_root / CONFIG_FILE).open("rb") as stream:
            config = tomllib.load(stream)
        directories = config.get("directories")
        if not isinstance(directories, dict):
            raise ValueError("config/paths.toml에 [directories] 설정이 필요합니다.")

        resolved: dict[str, Path] = {}
        for name in ("input", "output", "logs"):
            value = directories.get(name)
            if not isinstance(value, str) or not value.strip():
                raise ValueError(f"directories.{name}에 상대 경로를 지정하세요.")
            relative = Path(value)
            # 심볼릭 링크와 '..'를 해석한 실제 대상이 프로젝트 내부인지 검증한다.
            target = (project_root / relative).resolve()
            if (
                relative.is_absolute()
                or target == project_root
                or not target.is_relative_to(project_root)
            ):
                raise ValueError(
                    f"directories.{name}은 프로젝트 내부의 하위 경로여야 합니다."
                )
            for other_name, other_path in resolved.items():
                # 입력·출력·로그의 동일 경로와 상하위 관계를 모두 거부한다.
                if _paths_overlap(target, other_path):
                    raise ValueError(
                        f"directories.{name}과 directories.{other_name} 경로가 겹칩니다."
                    )
            resolved[name] = target

        return cls(
            root=project_root,
            input_dir=resolved["input"],
            output_dir=resolved["output"],
            logs_dir=resolved["logs"],
        )

    def list_input_files(self, *, recursive: bool = False) -> list[Path]:
        """입력 디렉토리에서 대상 엑셀 파일의 경로를 정렬해 반환한다.

        Args:
            recursive: True이면 하위 폴더도 탐색한다. 기본값 False는 직속 항목만
                확인한다. self.input_dir를 기준으로 파일 목록을 읽는다.

        Returns:
            대소문자와 관계없이 .xlsx 또는 .xlsb인 파일의 경로 목록.
            '~$'로 시작하는 파일은 제외한다. 전체 경로의 casefold() 결과로
            정렬하고 같으면 원래 경로 문자열로 비교한다. 대상이 없으면 빈 목록.

        Raises:
            FileNotFoundError: 입력 경로가 없거나 디렉토리로 확인되지 않는 경우.
            OSError: 디렉토리 탐색 중 파일시스템 접근 오류가 전파된 경우.

        Side Effects:
            디렉토리 항목과 파일 정보를 조회한다. 파일 내용 읽기와 쓰기는 없다.
            파일을 가리키는 심볼릭 링크도 포함할 수 있으며, 반환 경로의 대상이
            입력 디렉토리 안에 있는지까지 검증하지는 않는다.
        """
        if not self.input_dir.is_dir():
            raise FileNotFoundError(f"입력 디렉토리가 없습니다: {self.input_dir}")
        candidates = (
            self.input_dir.rglob("*") if recursive else self.input_dir.iterdir()
        )
        # Excel의 '~$' 임시 파일을 실제 파싱 입력에 포함하지 않도록 제외한다.
        files = [
            path
            for path in candidates
            if not path.name.startswith("~$")
            and path.suffix.lower() in WORKBOOK_EXTENSIONS
            and path.is_file()
        ]
        return sorted(files, key=lambda path: (str(path).casefold(), str(path)))
