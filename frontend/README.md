# frontend

Vue 3 + Vite + Vue Router + Pinia

## 실행

Node.js 20.19 이상이 필요합니다.

```bash
cd frontend
npm install
npm run dev      # http://localhost:5173
npm run build
```

개발 서버는 `/api` 요청을 `http://localhost:8080`(backend)으로 프록시합니다.

## 구조

```
src/
├── main.js        # 앱 진입점
├── App.vue
├── router/        # 라우트 정의
├── views/         # 페이지 단위 컴포넌트
├── components/    # 재사용 컴포넌트
├── stores/        # Pinia 스토어
└── assets/        # 정적 리소스
```
