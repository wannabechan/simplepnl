import { useState } from "react";
import { clearStoredToken, checkAccount, loginWithPassword, setStoredToken } from "./authClient";

const MSG_ACCESS = "アクセス不可";
const MSG_PASSWORD = "パスワードエラー";

type Props = {
  onSuccess: () => void;
};

export default function LoginScreen({ onSuccess }: Props) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [phase, setPhase] = useState<"account" | "password">("account");
  const [busy, setBusy] = useState(false);
  const [inlineError, setInlineError] = useState("");

  const handleCheckAccount = async () => {
    setInlineError("");
    const trimmed = email.trim();
    if (!trimmed) {
      setInlineError("이메일을 입력해 주세요.");
      return;
    }
    setBusy(true);
    try {
      const result = await checkAccount(trimmed);
      if (result.ok) {
        setPhase("password");
        setPassword("");
        return;
      }
      if (result.status === 403) {
        setInlineError(MSG_ACCESS);
        return;
      }
      if (result.status === 404) {
        setInlineError("인증 API를 찾을 수 없습니다. vercel dev로 실행했는지 확인해 주세요.");
        return;
      }
      if (result.status === 0 || Number.isNaN(result.status)) {
        setInlineError("서버에 연결할 수 없습니다. vercel dev로 실행했는지 확인해 주세요.");
        return;
      }
      setInlineError("확인에 실패했습니다. 잠시 후 다시 시도해 주세요.");
    } catch {
      setInlineError("네트워크 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const handleCheckPassword = async () => {
    setInlineError("");
    const trimmed = email.trim();
    if (!trimmed) {
      setInlineError("이메일을 입력해 주세요.");
      return;
    }
    if (!password) {
      setInlineError("비밀번호를 입력해 주세요.");
      return;
    }
    setBusy(true);
    try {
      const result = await loginWithPassword(trimmed, password);
      if (result.ok) {
        clearStoredToken();
        setStoredToken(result.token);
        onSuccess();
        return;
      }
      if (result.status === 403) {
        setInlineError(MSG_ACCESS);
        return;
      }
      if (result.status === 401) {
        setInlineError(MSG_PASSWORD);
        return;
      }
      if (result.status === 404) {
        setInlineError("인증 API를 찾을 수 없습니다. vercel dev로 실행했는지 확인해 주세요.");
        return;
      }
      setInlineError("로그인에 실패했습니다. 다시 시도해 주세요.");
    } catch {
      setInlineError("네트워크 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="login-wrap">
      <div className="panel login-card">
        <div className="panel-heading">
          <h2>login</h2>
        </div>
        <br />

        <input
          id="login-email"
          type="email"
          autoComplete="email"
          aria-label="이메일"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              if (phase === "account") void handleCheckAccount();
              else void handleCheckPassword();
            }
          }}
          placeholder="name@example.com"
          disabled={busy}
          style={{ width: "100%", marginBottom: 12 }}
        />

        {phase === "password" ? (
          <input
            id="login-password"
            type="password"
            autoComplete="current-password"
            aria-label="비밀번호"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") void handleCheckPassword();
            }}
            disabled={busy}
            style={{ width: "100%", marginBottom: 12 }}
          />
        ) : null}

        <button
          type="button"
          onClick={() => void (phase === "account" ? handleCheckAccount() : handleCheckPassword())}
          disabled={busy}
        >
          {busy ? "처리 중…" : phase === "account" ? "check account" : "check password"}
        </button>

        {inlineError ? (
          <p className="error" style={{ marginTop: 12, marginBottom: 0 }}>
            {inlineError}
          </p>
        ) : null}
      </div>
    </div>
  );
}
