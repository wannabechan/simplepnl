import { useState } from "react";
import { clearStoredToken, checkAccount, loginWithPassword, setStoredToken } from "./authClient";

const fmtAccountErr = (code: string | number) => `アクセス不可 : ${code}`;
const fmtPasswordErr = (code: string | number) => `パスワードエラー : ${code}`;

function httpOrNetworkCode(status: number): string | number {
  if (!Number.isFinite(status) || status === 0) return "NETWORK";
  return status;
}

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
      setInlineError(fmtAccountErr("INPUT"));
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
      setInlineError(fmtAccountErr(httpOrNetworkCode(result.status)));
    } catch {
      setInlineError(fmtAccountErr("NETWORK"));
    } finally {
      setBusy(false);
    }
  };

  const handleCheckPassword = async () => {
    setInlineError("");
    const trimmed = email.trim();
    if (!trimmed) {
      setInlineError(fmtPasswordErr("INPUT_EMAIL"));
      return;
    }
    if (!password) {
      setInlineError(fmtPasswordErr("INPUT_PASSWORD"));
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
      setInlineError(fmtPasswordErr(httpOrNetworkCode(result.status)));
    } catch {
      setInlineError(fmtPasswordErr("NETWORK"));
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
