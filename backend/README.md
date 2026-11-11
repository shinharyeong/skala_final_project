# backend

Spring Boot 3.4 · Java 21 · Gradle (Kotlin DSL)

## 실행

JDK 21이 필요합니다.

```bash
cd backend
./gradlew bootRun    # http://localhost:8080
./gradlew test
```

상태 확인: `GET http://localhost:8080/actuator/health`

## 구조

```
src/main/java/io/skala/backend/   # 애플리케이션 코드
src/main/resources/application.yml
src/test/java/io/skala/backend/   # 테스트
```

로컬 전용 설정(DB 비밀번호 등)은 `application-local.yml`에 두고 커밋하지 않습니다.
