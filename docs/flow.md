# 업무 흐름도

[도메인 규칙](domain-rules.md)을 단계별 흐름으로 나타낸 문서입니다.

**색 구분**

| 색 | 주체 |
| --- | --- |
| 파랑 | 경영기획 담당자 |
| 노랑 | 조직 담당자 |
| 회색 | 시스템 (자동 처리) |
| 점선 | 미정 사항 |

## 전체 단계

```mermaid
flowchart LR
    S1["1. 연간 계획 수립"] --> S2["2. 월 실적 제출"]
    S2 --> S3["3. Gap 검토·소명"]
    S3 --> S4["4. master·대시보드·보고"]
    S4 -->|다음 달| S2
```

## 1. 연간 계획 수립

```mermaid
flowchart TD
    A(["시작"])
    B["연간 계획 제출 요청<br/>(대상 연도 지정)"]
    C["조직 담당자에게 알림"]
    D["연간 계획 엑셀 업로드"]
    E{"파일 형식·필수값<br/>자동 검증 통과?"}
    E1["오류 위치 표시<br/>(제출되지 않음)"]
    F["제출본 생성<br/>상태: 검토 대기"]
    G{"계획이 올바른가?"}
    H["재제출 요청 + 사유"]
    I["재제출 요청 알림"]
    J["연간 계획 수정"]
    K["조직 계획 승인"]
    L{"모든 조직<br/>승인 완료?"}
    M["남은 조직 대기"]
    N["전사 계획 합산"]
    O["Contingency 입력<br/>(전사 기준, 1회)"]
    P["연간 계획 확정·잠금<br/>보고 영업이익 = 합산 − Contingency"]
    Q(["월 실적 단계로"])

    A --> B --> C --> D --> E
    E -->|아니오| E1 --> D
    E -->|예| F --> G
    G -->|아니오| H --> I --> J --> D
    G -->|예| K --> L
    L -->|아니오| M
    L -->|예| N --> O --> P --> Q

    classDef planner fill:#E3EDFD,stroke:#3B6FD8,color:#1A2B4C
    classDef org fill:#FFF3D1,stroke:#C99400,color:#3D2E00
    classDef system fill:#EFEFEF,stroke:#8A8A8A,color:#222
    classDef terminal fill:#D7F4DD,stroke:#3E9E55,color:#173D20
    class B,G,H,K,O planner
    class D,J org
    class C,E,E1,F,I,L,M,N,P system
    class A,Q terminal
```

## 2. 월 실적 제출

```mermaid
flowchart TD
    A(["월 실적 시즌"])
    B["월 실적 제출 요청<br/>(대상 월 지정)"]
    C["조직 담당자에게 알림"]
    D{"업로드 가능?<br/>열린 요청 있음<br/>그리고 검토 대기 제출본 없음"}
    D1["업로드 버튼 비활성<br/>(사유 표시)"]
    E["실적 + 다음 달 추정치<br/>엑셀 업로드"]
    F{"파서 자동 검증 통과?"}
    F1["오류 위치 표시<br/>(제출되지 않음)"]
    G["마지막 승인본과 비교<br/>(파일에 든 모든 월)"]
    H{"이미 제출한 값이<br/>바뀌었나?"}
    I["변경 항목 표시<br/>예) 2월 광고비 1,200 → 1,350"]
    J["항목별 변경 사유 입력<br/>(필수)"]
    K["제출 완료<br/>버전 YYYY-MM-NN 부여<br/>상태: 검토 대기"]
    L["경영기획 담당자에게 알림<br/>(변경 사유 포함)"]
    M(["Gap 검토 단계로"])

    A --> B --> C --> D
    D -->|아니오| D1
    D -->|예| E --> F
    F -->|아니오| F1 --> E
    F -->|예| G --> H
    H -->|예| I --> J --> K
    H -->|아니오| K
    K --> L --> M

    classDef planner fill:#E3EDFD,stroke:#3B6FD8,color:#1A2B4C
    classDef org fill:#FFF3D1,stroke:#C99400,color:#3D2E00
    classDef system fill:#EFEFEF,stroke:#8A8A8A,color:#222
    classDef terminal fill:#D7F4DD,stroke:#3E9E55,color:#173D20
    class B planner
    class E,J org
    class C,D,D1,F,F1,G,H,I,K,L system
    class A,M terminal
```

## 3. Gap 검토·소명

```mermaid
flowchart TD
    A(["검토 대기 제출본"])
    B["Gap 계산<br/>조직 × 항목 × 월 손익<br/>실적 − 계획"]
    C["기준 슬라이더 조정<br/>(% · 금액, 실시간 미리보기)"]
    D{"적용?"}
    E["기준값 저장"]
    F["기준 초과 Gap 목록<br/>% 그리고 금액 초과<br/>+ 계획 0, 실적 ≠ 0"]
    G{"항목별<br/>산출 근거 명확?"}
    H["확인(인정)"]
    I["Gap 소명 요청"]
    J["소명 요청 알림"]
    K["소명 작성"]
    L{"소명이 적절한가?"}
    L1["소명 재요청"]
    M["소명 확정<br/>소명 모음·조직별 현황에 저장"]
    N{"모든 Gap 항목<br/>처리 완료?"}
    V["변경 사유 확인<br/>(있는 경우)"]
    O{"제출본 승인?"}
    P["승인·아카이브<br/>조직별 현황에 버전 기록"]
    R["재제출 요청"]
    R1["상태: 재제출 요청<br/>다음 회차 업로드 가능"]
    Q(["master 단계로"])

    A --> B --> C --> D
    D -->|아니오| C
    D -->|예| E --> F --> G
    G -->|예| H --> N
    G -->|아니오| I --> J --> K --> L
    L -->|아니오| L1 --> K
    L -->|예| M --> N
    N -->|아니오| G
    N -->|예| V --> O
    O -->|예| P --> Q
    O -->|아니오| R --> R1

    classDef planner fill:#E3EDFD,stroke:#3B6FD8,color:#1A2B4C
    classDef org fill:#FFF3D1,stroke:#C99400,color:#3D2E00
    classDef system fill:#EFEFEF,stroke:#8A8A8A,color:#222
    classDef terminal fill:#D7F4DD,stroke:#3E9E55,color:#173D20
    class C,D,G,H,I,L,L1,V,O,R planner
    class K org
    class B,E,F,J,M,N,P,R1 system
    class A,Q terminal
```

## 4. master·대시보드·보고

```mermaid
flowchart TD
    A(["조직 제출본 승인"])
    B{"해당 월 모든 조직<br/>승인 완료?"}
    B1["대기<br/>(최초 생성 시점·미제출 조직 처리 미정)"]
    C["master 생성<br/>master-YYYY-MM-01<br/>조직 승인본 합산"]
    D["홈: 전사 손익 현황 갱신<br/>합산 영업이익 / Contingency / 보고 영업이익"]
    E["연·월 필터로 이전 master 조회<br/>소명 요청·회신 현황"]
    F["보고서 화면<br/>전월·전년 대비 표, 주별 차트, 소명 메모<br/>사용한 master 버전 기록"]
    G["표·차트 내보내기"]
    H(["끝: 각자 장표에 활용"])
    X{"승인본이 이미 master가 있는<br/>월의 값을 바꿨나?<br/>(같은 달 재제출 · 과거 월 변경)"}
    Y["해당 월 master 새 버전 생성<br/>master-YYYY-MM-NN<br/>이전 버전 보존"]

    A --> X
    X -->|예| Y --> D
    X -->|아니오| B
    B -->|아니오| B1
    B -->|예| C --> D --> E
    D --> F --> G --> H

    classDef planner fill:#E3EDFD,stroke:#3B6FD8,color:#1A2B4C
    classDef system fill:#EFEFEF,stroke:#8A8A8A,color:#222
    classDef terminal fill:#D7F4DD,stroke:#3E9E55,color:#173D20
    classDef undecided fill:#FFFFFF,stroke:#8A8A8A,stroke-dasharray:5 4,color:#555
    class E,F,G planner
    class B,C,D,X,Y system
    class A,H terminal
    class B1 undecided
```

## 상태

### 제출본

```mermaid
stateDiagram-v2
    state "검토 대기" as pending
    state "재제출 요청" as rejected
    state "승인" as approved

    [*] --> pending: 업로드·제출 완료
    pending --> approved: 모든 Gap 처리 후 승인
    pending --> rejected: 반려
    rejected --> [*]: 다음 회차 업로드 가능
    approved --> [*]: 아카이브
```

### Gap 항목

```mermaid
stateDiagram-v2
    state "미검토" as open
    state "확인(인정)" as accepted
    state "소명 요청" as requested
    state "소명 작성됨" as answered
    state "소명 확정" as resolved

    [*] --> open: 기준 초과
    open --> accepted: 산출 근거 명확
    open --> requested: 산출 근거 불명확
    requested --> answered: 조직 담당자 작성
    answered --> resolved: 적절
    answered --> requested: 재요청
    accepted --> [*]
    resolved --> [*]
```
