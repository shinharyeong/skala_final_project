# Git 협업 컨벤션

형상 관리와 협업 규칙입니다. 충돌을 줄이고 `main`을 항상 배포 가능한 상태로 유지하는 것이 목표입니다.

## 1. 폴더 구조

모노레포로 관리하며, 영역별 코드는 각 폴더에만 둡니다.

```
root/
├── frontend/      # 프론트엔드
├── backend/       # Spring Boot 백엔드
├── ai/            # FastAPI·LangChain 등 AI 기능
├── parser/        # Excel 파싱
├── docs/          # 문서
├── .github/       # 이슈·PR 템플릿, 워크플로
├── .gitignore     # 공통 설정
└── README.md
```

각자 저장소 전체를 clone하되, 실행과 IDE는 자기 영역 폴더 기준으로 합니다.

## 2. 브랜치 전략

| 브랜치 | 이름 | 역할 | 분기 → 병합 |
| --- | --- | --- | --- |
| 배포 | `main` | 언제나 실행·배포 가능한 코드 | — |
| 통합 | `develop` | 다음 배포에 들어갈 기능이 모이는 곳 (기본 브랜치) | — |
| 기능 | `feat/{scope}-{설명}` | 기능 개발 | `develop` → `develop` |
| 수정 | `fix/{scope}-{설명}` | `develop`에서 발견한 버그 수정 | `develop` → `develop` |
| 긴급 수정 | `hotfix/{scope}-{설명}` | `main`(배포본)의 긴급 버그 수정 | `main` → `main`, `develop` |
| 기타 | `refactor/…` `docs/…` `chore/…` `test/…` | 리팩터링·문서·설정·테스트 | `develop` → `develop` |

**규칙**

- `main`, `develop`에는 직접 push하지 않습니다. 모든 변경은 PR로 합칩니다.
- `main`에는 `develop`과 `hotfix/*`에서 온 PR만 병합합니다.
- `hotfix`는 `main`에 병합한 뒤 같은 변경을 `develop`에도 반영합니다.
- 브랜치 이름은 영문 소문자와 하이픈만 씁니다. 병합한 브랜치는 삭제합니다.

```
feat/fe-login-page
feat/be-auth-api
fix/parser-master-compare-tolerance
hotfix/be-token-expiry
docs/git-convention
```

## 3. 커밋 메시지

**형식:** `<type>(<scope>): <subject>`

```
feat(fe): 로그인 페이지 UI 구현
fix(be): JWT 토큰 만료 시간 수정
chore(common): 루트 .gitignore 정리
```

- subject는 한국어로, 마침표 없이 50자 안쪽으로 씁니다.
- 필요하면 한 줄 띄우고 본문에 무엇을·왜 바꿨는지 씁니다.
- 관련 이슈는 본문 끝에 `#12`처럼 적습니다.

### Type

| type | 설명 |
| --- | --- |
| `feat` | 새 기능 |
| `fix` | 버그 수정 |
| `refactor` | 기능 변경 없는 코드 개선 |
| `style` | 포맷팅 등 로직 변경 없는 수정 |
| `docs` | 문서·주석 수정 (README, Swagger 등) |
| `test` | 테스트 추가·수정 |
| `chore` | 빌드·패키지·설정 등 기타 |

### Scope

| scope | 대상 |
| --- | --- |
| `fe` | `frontend/` |
| `be` | `backend/` |
| `ai` | `ai/` |
| `parser` | `parser/` |
| `docs` | `docs/` |
| `common` | 루트 설정(`.gitignore`, `README.md`, `.github/`) 또는 여러 영역에 걸친 변경 |

## 4. 이슈

기록이나 논의가 필요한 작업은 이슈부터 만듭니다. 작은 작업은 이슈 없이 바로 PR을 올려도 됩니다.

| 이슈 먼저 | 바로 PR |
| --- | --- |
| 기능 개발, 버그, 리뷰 지적사항 | 오타·문구 수정 |
| 여러 PR에 걸치는 작업 | 작은 설정·의존성 변경 |
| 방향을 먼저 논의해야 하는 작업 | PR 설명만으로 충분한 작은 수정 |

- **제목:** `[type] 요약` — 예: `[feat] 로그인 API 구현`, `[fix] 마스터 비교 오차`
- 템플릿(버그 / 기능 / 작업)을 고르고 영역을 선택합니다.
- 담당자(Assignees)와 라벨을 지정합니다. 진행 상황은 GitHub Projects 보드로 관리합니다.

**라벨**

| 라벨 | 용도 |
| --- | --- |
| `bug` / `enhancement` / `documentation` / `question` | 종류 |
| `fe` / `be` / `ai` / `parser` | 영역 |

## 5. Pull Request

**제목:** 커밋과 같은 형식 `<type>(<scope>): <subject>` — 예: `feat(be): 회원가입 API 구현`

Squash merge 시 PR 제목이 그대로 `develop`의 커밋 메시지가 되므로 형식을 지킵니다.

- 기본 대상은 `develop`입니다. (`hotfix`만 `main`)
- PR 템플릿을 채웁니다. 해당 없는 섹션은 지우지 말고 "해당 없음"이라고 적습니다.
- 관련 이슈가 있으면 `Closes #12`로 연결하고, 없으면 이유를 적습니다.
- **실제로 실행한 검증만** 적고, 확인하지 못한 부분은 "검증하지 못한 부분"에 남깁니다.
- AI 도구로 작성한 코드도 작성자가 직접 읽고 이해한 뒤 올립니다. PR 전체의 책임은 작성자에게 있습니다.
- 리뷰어 1명 이상 지정, **승인 1개 이상** 받은 뒤 병합합니다.
- 한 PR에는 가능하면 한 영역만 담습니다. API 변경처럼 여러 영역을 함께 고쳐야 하면 scope를 `common`으로 하고 본문에 이유를 적습니다.
- 입력 데이터, 실행 결과, `.env` 등 비밀 정보는 커밋하지 않습니다.

**병합 방식**

| 방향 | 방식 |
| --- | --- |
| 작업 브랜치 → `develop` | Squash and merge |
| `develop` → `main` (배포) | Create a merge commit |
| `hotfix/*` → `main` | Create a merge commit, 이후 `develop`에 반영 |

**리뷰 코멘트 접두어**

| 접두어 | 의미 |
| --- | --- |
| `[필수]` | 병합 전 반드시 수정 |
| `[제안]` | 반영 여부는 작성자 판단 |
| `[질문]` | 의도 확인 |

## 6. 작업 순서

```bash
# 1. (필요 시) 이슈 생성 — 예: [feat] 로그인 API 구현 (#12)

# 2. develop 최신화 후 브랜치 생성
git switch develop
git pull origin develop
git switch -c feat/be-login

# 3. 자기 영역 폴더에서 개발, 규칙에 맞춰 커밋
git commit -m "feat(be): 로그인 로직 구현"

# 4. push 후 develop으로 PR 생성
git push -u origin feat/be-login

# 5. 리뷰 승인 → Squash and merge → 브랜치 삭제
```

작업 중 `develop`이 많이 바뀌었으면 `git pull --rebase origin develop`으로 최신화한 뒤 push합니다.

## 7. 저장소 설정 (관리자)

- 기본 브랜치: `develop`
- Branch protection (`main`, `develop`): PR 필수, 승인 1개 이상, force push 금지
- Merge 옵션: Squash merge·Merge commit 허용, 병합 후 브랜치 자동 삭제
- CI가 추가되면 통과를 병합 조건에 포함합니다.
