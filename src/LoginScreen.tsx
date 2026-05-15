import { useState } from "react";
import { clearStoredToken, requestOtp, setStoredToken, verifyOtp } from "./authClient";

type Props = {
  onSuccess: () => void;
};

export default function LoginScreen({ onSuccess }: Props) {
  const [email, setEmail] = useState("");
  const [code, setCode] = useState("");
  const [phase, setPhase] = useState<"email" | "code">("email");
  const [busy, setBusy] = useState(false);
  const [inlineError, setInlineError] = useState("");

  const handleRequestOtp = async () => {
    setInlineError("");
    const trimmed = email.trim();
    if (!trimmed) {
      setInlineError("이메일을 입력해 주세요.");
      return;
    }
    setBusy(true);
    try {
      const result = await requestOtp(trimmed);
      if (result.ok) {
        setPhase("code");
        setCode("");
        return;
      }
      if (result.status === 403) {
        window.alert("접근이 불가합니다.");
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
      setInlineError("코드 발송에 실패했습니다. 잠시 후 다시 시도해 주세요.");
    } catch {
      setInlineError("네트워크 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const handleVerify = async () => {
    setInlineError("");
    const digits = code.replace(/\D/g, "").slice(0, 6);
    if (digits.length !== 6) {
      setInlineError("6자리 코드를 입력해 주세요.");
      return;
    }
    setBusy(true);
    try {
      const result = await verifyOtp(email.trim(), digits);
      if (result.ok) {
        clearStoredToken();
        setStoredToken(result.token);
        onSuccess();
        return;
      }
      if (result.status === 401) {
        setInlineError("코드가 올바르지 않거나 만료되었습니다.");
        return;
      }
      setInlineError("로그인에 실패했습니다. 다시 시도해 주세요.");
    } catch {
      setInlineError("네트워크 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const goBackToEmail = () => {
    setPhase("email");
    setCode("");
    setInlineError("");
  };

  return (
    <div className="login-wrap">
      <div className="panel login-card">
        <div className="panel-heading">
          <h2>로그인</h2>
        </div>

        {phase === "email" ? (
          <>
            <input
              id="login-email"
              type="email"
              autoComplete="email"
              aria-label="이메일"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") void handleRequestOtp();
              }}
              placeholder="name@example.com"
              disabled={busy}
              style={{ width: "100%", marginBottom: 12 }}
            />
            <button type="button" onClick={() => void handleRequestOtp()} disabled={busy}>
              {busy ? "처리 중…" : "인증 코드 발송"}
            </button>
          </>
        ) : (
          <>
            <p className="muted" style={{ fontSize: 14 }}>
              <strong>{email.trim()}</strong> 로 발송된 6자리 코드를 입력해 주세요.
            </p>
            <label className="login-label" htmlFor="login-code">
              인증 코드
            </label>
            <input
              id="login-code"
              type="text"
              inputMode="numeric"
              autoComplete="one-time-code"
              maxLength={6}
              value={code}
              onChange={(e) => setCode(e.target.value.replace(/\D/g, "").slice(0, 6))}
              onKeyDown={(e) => {
                if (e.key === "Enter") void handleVerify();
              }}
              placeholder="000000"
              disabled={busy}
              style={{ width: "100%", marginBottom: 12, letterSpacing: "0.2em", fontSize: 18 }}
            />
            <div className="login-actions">
              <button type="button" onClick={() => void handleVerify()} disabled={busy}>
                {busy ? "처리 중…" : "로그인"}
              </button>
              <button type="button" className="login-secondary" onClick={goBackToEmail} disabled={busy}>
                이메일 변경
              </button>
            </div>
          </>
        )}

        {inlineError ? (
          <p className="error" style={{ marginTop: 12, marginBottom: 0 }}>
            {inlineError}
          </p>
        ) : null}
      </div>
    </div>
  );
}
