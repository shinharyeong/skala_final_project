# ERD 초안

```mermaid
erDiagram
    organization ||--o{ app_user : "소속"
    organization ||--o{ submission : "제출"
    annual_plan ||--o{ submission_request : "연도"
    app_user ||--o{ submission_request : "요청"
    submission_request ||--o{ submission : "요청에 대한 제출"
    submission ||--o{ financial_value : "파싱 결과"
    account_item ||--o{ financial_value : "항목"
    account_item ||--o| account_item : "상위 항목"
    submission ||--o{ change_reason : "변경 사유"
    account_item ||--o{ change_reason : "항목"
    submission ||--o{ gap_item : "Gap"
    account_item ||--o{ gap_item : "항목"
    gap_threshold ||--o{ gap_item : "적용 기준"
    gap_item ||--o{ gap_explanation : "소명 회차"
    master_version ||--|{ master_source : "구성"
    submission ||--o{ master_source : "포함"
    app_user ||--o{ notification_event : "수신"

    organization {
        bigint id PK
        varchar code UK "사업부코드"
        varchar name "조직명"
        boolean active
    }

    app_user {
        bigint id PK
        varchar email UK
        varchar password_hash
        varchar name
        varchar role "PLANNER | ORG_MANAGER"
        bigint organization_id FK "ORG_MANAGER만, PLANNER는 null"
    }

    annual_plan {
        bigint id PK
        int year UK
        varchar status "COLLECTING | CONFIRMED"
        numeric contingency_amount "전사 단위, 확정 시 잠금"
        bigint confirmed_by FK
        timestamp confirmed_at
    }

    submission_request {
        bigint id PK
        bigint annual_plan_id FK
        varchar type "ANNUAL_PLAN | MONTHLY_ACTUAL"
        int year
        int month "ANNUAL_PLAN이면 null"
        varchar status "OPEN | CLOSED"
        date due_date "마감 규칙 미정"
        bigint requested_by FK
        timestamp requested_at
    }

    submission {
        bigint id PK
        bigint request_id FK
        bigint organization_id FK
        int seq "회차 NN"
        varchar version_label UK "YYYY-MM-NN (조직별)"
        varchar status "PENDING | REJECTED | APPROVED"
        varchar s3_key "원본 엑셀"
        char sha256
        bigint submitted_by FK
        timestamp submitted_at
        bigint reviewed_by FK
        timestamp reviewed_at
        text reject_reason
    }

    account_item {
        bigint id PK
        varchar code UK "원본 항목·예산코드"
        varchar name
        bigint parent_id FK
    }

    financial_value {
        bigint id PK
        bigint submission_id FK
        bigint item_id FK
        int year
        int month
        varchar scenario "PLAN | ACTUAL | ESTIMATE"
        varchar measure "REVENUE | COST"
        numeric amount "원 단위, Decimal"
        varchar source_cell "시트!셀"
    }

    change_reason {
        bigint id PK
        bigint submission_id FK
        bigint item_id FK
        int year
        int month
        varchar measure
        numeric previous_value "마지막 승인본 값"
        numeric new_value
        text reason
    }

    gap_threshold {
        bigint id PK
        numeric percent
        numeric amount
        bigint updated_by FK
        timestamp applied_at "적용 버튼 시점, 이력 보존"
    }

    gap_item {
        bigint id PK
        bigint submission_id FK
        bigint organization_id FK
        bigint item_id FK
        int year
        int month
        numeric plan_profit "계획 손익"
        numeric actual_profit "실적 손익"
        numeric gap_amount
        numeric gap_rate "계획 0이면 null"
        bigint threshold_id FK
        varchar status "OPEN | ACCEPTED | REQUESTED | ANSWERED | RESOLVED"
    }

    gap_explanation {
        bigint id PK
        bigint gap_item_id FK
        int round "요청 회차"
        text request_note
        bigint requested_by FK
        timestamp requested_at
        text answer
        bigint answered_by FK
        timestamp answered_at
        varchar verdict "PENDING | ACCEPTED | REJECTED"
    }

    master_version {
        bigint id PK
        int year
        int month
        int seq
        varchar version_label UK "master-YYYY-MM-NN"
        bigint trigger_submission_id FK "갱신을 일으킨 승인본"
        timestamp created_at
    }

    master_source {
        bigint master_version_id PK,FK
        bigint organization_id PK,FK
        bigint submission_id FK "그 시점 조직 승인본"
    }

    notification_event {
        bigint id PK
        bigint recipient_id FK
        varchar event_type
        jsonb payload
        timestamp created_at
        timestamp read_at
    }
```
