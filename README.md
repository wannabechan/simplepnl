# simplepnl

**월·매장 구성**만 다루는 웹앱입니다. (손익·파일 업로드·비용 정리 등은 제거된 상태이며, 이후 다시 설계 예정입니다.)

## 1) Vercel Storage (Postgres/Neon) 준비

1. Vercel 프로젝트 > Storage 탭에서 Postgres(Neon) 연동
2. 프로젝트 환경변수에 `DATABASE_URL`이 자동 추가되는지 확인

## 2) 로컬 환경변수

`.env` 파일을 만들고 아래 값 입력:

```bash
DATABASE_URL=...
ACCOUNT_ADMIN=관리자@이메일.com
RESEND_API_KEY=re_...
# 선택
RESEND_FROM_EMAIL=SimplePNL <notify@yourdomain.com>
```

> 앱 상태는 **서버(DB)만** 사용합니다. 브라우저 localStorage에는 저장하지 않습니다.
> **첫 화면은 로그인**입니다. `ACCOUNT_ADMIN`과 동일한 이메일만 인증 코드 발송이 허용되며, 코드는 **Resend**로 발송됩니다. 관리자 이메일 비교·코드 검증은 **서버(API)** 에서만 수행됩니다.
> 로컬에서 배포와 동일한 DB·API를 보려면 `vercel link` / `vercel env pull`로 환경변수를 맞춘 뒤 **`vercel dev`** 로 실행하세요. (`npm run dev`만 쓰면 `/api/*`가 없어 로그인·불러오기/저장이 실패합니다.)

`vercel dev`에서 **`startsWith` / exit code 1** 이 나오던 경우는, Vite 루트의 solution형 `tsconfig.json`과 맞물려 **`api/*.ts` 서버리스 번들이 깨지는** CLI 이슈로 보입니다. 그래서 API는 **`api/*.mjs`**(순수 ESM, TypeScript 미사용)로 두어 해당 경로를 피했습니다. 변경 후 **`vercel dev`를 완전히 종료했다가 다시 실행**하세요.

`@neondatabase/serverless`는 **Node 19+** 엔진을 요구합니다. 로컬 Node가 너무 낮으면 `nvm use 20`(또는 22) 등으로 올린 뒤 다시 시도하세요.

## 3) 실행

```bash
npm install
vercel dev
```

UI만 빠르게 볼 때(저장 없이):

```bash
npm run dev
```

(`npm run dev`만으로는 로그인·`/api/state`가 동작하지 않습니다.)

## 4) Vercel 환경변수

Vercel Project Settings > Environment Variables에서 아래 값 확인:

- `DATABASE_URL`
- `ACCOUNT_ADMIN` — 허용할 관리자 이메일 한 주소
- `RESEND_API_KEY` — 인증 메일 발송용 ([Resend](https://resend.com)에서 발급)
- `RESEND_FROM_EMAIL` (선택) — 발신 주소. 미설정 시 Resend 기본 테스트 발신 주소를 사용합니다.

배포 후 앱은 `/api/state`를 통해 DB에만 저장합니다. 로그인 세션 토큰은 브라우저 **sessionStorage**에만 보관됩니다(탭을 닫으면 다시 로그인).
