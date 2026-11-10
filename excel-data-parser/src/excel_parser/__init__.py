"""엑셀 데이터 파서의 공개 경로 관리 인터페이스를 제공한다.

외부 호출자는 내부 모듈 경로 대신 이 패키지에서 ProjectPaths를 가져온다.
"""

from .paths import ProjectPaths

__all__ = ["ProjectPaths"]
