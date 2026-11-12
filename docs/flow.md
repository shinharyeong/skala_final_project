# 업무 흐름도

[도메인 규칙](domain-rules.md)을 단계별 흐름으로 나타낸 문서입니다.

**색 구분**

| 색 | 주체 |
| --- | --- |
| 파랑 | 경영기획 담당자 |
| 노랑 | 조직 담당자 |
| 회색 | 시스템 (자동 처리) |

## 전체 단계

```mermaid
flowchart LR
    S1["1. 연간 계획 수립"] --> S2["2. 월 실적 제출"]
    S2 --> S3["3. 소명 검토·승인"]
    S3 --> S4["4. master·대시보드·보고"]
    S4 -->|다음 달| S2
```

## 1. 연간 계획 수립

```mermaid
flowchart TD
    A(["시작"])
    B["연간 계획 제출 요청<br/>(대상 연도 지정)"]
    C["조직 담당자에게 알림"]
    D["연간 계획 엑셀 업로드<br/>예) D100_2026-PLAN.xlsx"]
    E{"파일 이름 규칙·형식·필수값<br/>자동 검증 통과?"}
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
    E["엑셀 업로드<br/>해당 월 실적 + 다음 달~12월 추정<br/>예) D100_2026-02.xlsx"]
    FN{"파일 이름 규칙?<br/>조직코드 · 대상 월 일치"}
    FN1["이름 오류 표시<br/>(파싱하지 않음)"]
    F{"파서 자동 검증 통과?"}
    F1["오류 위치 표시<br/>(제출되지 않음)"]
    GP["Gap 계산 (칸별)<br/>계획 대비: 해당 월 실적 vs 연간 계획<br/>직전 대비: 모든 칸 vs 직전 버전"]
    X{"기준 초과 칸?<br/>실적·추정 기준별<br/>% 그리고 금액"}
    Y["기준 초과 칸 표시<br/>수용된 소명 중 값이 같은 칸은 이어받기"]
    Z["칸마다 소명 하나 작성 (필수)<br/>태그: 버전·월·항목·실적/추정·[계획 대비][직전 대비]"]
    K["제출 완료<br/>저장 이름 D100_2026-02-NN.xlsx<br/>상태: 검토 대기"]
    L["경영기획 담당자에게 알림<br/>(소명 포함)"]
    M(["소명 검토 단계로"])

    A --> B --> C --> D
    D -->|아니오| D1
    D -->|예| E --> FN
    FN -->|아니오| FN1 --> E
    FN -->|예| F
    F -->|아니오| F1 --> E
    F -->|예| GP --> X
    X -->|예| Y --> Z --> K
    X -->|아니오| K
    K --> L --> M

    classDef planner fill:#E3EDFD,stroke:#3B6FD8,color:#1A2B4C
    classDef org fill:#FFF3D1,stroke:#C99400,color:#3D2E00
    classDef system fill:#EFEFEF,stroke:#8A8A8A,color:#222
    classDef terminal fill:#D7F4DD,stroke:#3E9E55,color:#173D20
    class B planner
    class E,Z org
    class C,D,D1,FN,FN1,F,F1,GP,X,Y,K,L system
    class A,M terminal
```

## 3. 소명 검토·승인

```mermaid
flowchart TD
    A(["검토 대기 제출본"])
    B["소명 목록<br/>(버전·월·항목·비교 태그)"]
    G{"소명이 적절한가?<br/>(칸별)"}
    H["수용"]
    I["소명 재요청<br/>(재요청 사유 첨부, 버전 그대로)"]
    J["소명 재요청 알림"]
    K["그 소명만 다시 작성<br/>(파일 재업로드 없음)"]
    N{"모든 소명<br/>수용 완료?"}
    O{"제출본 승인?"}
    P["승인·아카이브<br/>조직별 현황에 버전·소명 기록"]
    R["재제출 요청"]
    R1["상태: 재제출 요청<br/>그 요청에 대한 재제출만 가능"]
    Q(["master 단계로"])

    A --> B --> G
    G -->|예| H --> N
    G -->|아니오| I --> J --> K --> G
    N -->|아니오| G
    N -->|예| O
    O -->|예| P --> Q
    O -->|아니오| R --> R1

    classDef planner fill:#E3EDFD,stroke:#3B6FD8,color:#1A2B4C
    classDef org fill:#FFF3D1,stroke:#C99400,color:#3D2E00
    classDef system fill:#EFEFEF,stroke:#8A8A8A,color:#222
    classDef terminal fill:#D7F4DD,stroke:#3E9E55,color:#173D20
    class B,G,H,I,O,R planner
    class K org
    class J,N,P,R1 system
    class A,Q terminal
```

### Gap 기준 설정 (경영기획 담당자, 수시)

```mermaid
flowchart LR
    T1["기준 선택<br/>실적 칸 / 추정 칸"]
    T2["슬라이더 조정<br/>% · 금액, 실시간 미리보기"]
    T3{"적용?"}
    T4["기준값 저장<br/>이후 제출부터 적용"]

    T1 --> T2 --> T3
    T3 -->|아니오| T2
    T3 -->|예| T4

    classDef planner fill:#E3EDFD,stroke:#3B6FD8,color:#1A2B4C
    classDef org fill:#FFF3D1,stroke:#C99400,color:#3D2E00
    classDef system fill:#EFEFEF,stroke:#8A8A8A,color:#222
    classDef terminal fill:#D7F4DD,stroke:#3E9E55,color:#173D20
    class T1,T2,T3 planner
    class T4 system
```

## 4. master·대시보드·보고

```mermaid
flowchart TD
    A(["조직 제출본 승인"])
    B{"해당 월 모든 조직<br/>승인 완료?"}
    B1["대기<br/>홈에 master 미생성 (승인 N/M 조직)<br/>미승인 조직 목록 표시"]
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
    class E,F,G planner
    class B,B1,C,D,X,Y system
    class A,H terminal
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

### Gap 항목 (소명)

```mermaid
stateDiagram-v2
    state "소명 필요" as required
    state "소명 작성됨" as answered
    state "재요청" as rerequested
    state "수용" as accepted

    [*] --> required: 제출 시 기준 초과
    [*] --> accepted: 값이 같아 수용된 소명 이어받기
    required --> answered: 조직이 소명 작성 후 제출
    answered --> accepted: 경영기획 수용
    answered --> rerequested: 경영기획 재요청
    rerequested --> answered: 조직이 다시 작성
    accepted --> [*]
```
