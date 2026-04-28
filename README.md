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

> 앱 상태는 **서버(DB)만** 사용합니다. 브라우저 localStorage에는 저장하지 않습니다.
> 로컬에서 배포와 동일한 DB를 보려면 `vercel link` / `vercel env pull`로 `DATABASE_URL`을 맞춘 뒤 **`vercel dev`** 로 실행하세요. (`npm run dev`만 쓰면 `/api/state`가 없어 불러오기/저장이 실패합니다.)

## 3) 실행

```bash
npm install
vercel dev
```

UI만 빠르게 볼 때(저장 없이):

```bash
npm run dev
```

## 4) Vercel 환경변수

Vercel Project Settings > Environment Variables에서 아래 값 확인:

- `DATABASE_URL`

배포 후 앱은 `/api/state`를 통해 DB에만 저장합니다.
