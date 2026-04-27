# simplepnl

월별 매장 손익 리포트 웹앱입니다.

## 1) Vercel Storage (Postgres/Neon) 준비

1. Vercel 프로젝트 > Storage 탭에서 Postgres(Neon) 연동
2. 프로젝트 환경변수에 `DATABASE_URL`이 자동 추가되는지 확인

## 2) 로컬 환경변수

`.env` 파일을 만들고 아래 값 입력:

```bash
DATABASE_URL=...
```

> 로컬에서 `npm run dev` 실행 시 `/api`는 Vite에서 동작하지 않으므로 localStorage 캐시로 동작합니다.
> 서버 저장까지 로컬에서 테스트하려면 `vercel dev`를 사용하세요.

## 3) 실행

```bash
npm install
npm run dev
```

## 4) Vercel 환경변수

Vercel Project Settings > Environment Variables에서 아래 값 확인:

- `DATABASE_URL`

배포 후 앱은 `/api/state`를 통해 DB에 저장하고, 브라우저에는 캐시로도 저장합니다.
브라우저 데이터 삭제 후에도 DB 데이터는 유지됩니다.
