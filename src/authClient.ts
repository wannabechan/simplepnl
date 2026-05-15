const TOKEN_KEY = "simplepnl_session_token";

export function getStoredToken(): string | null {
  try {
    return sessionStorage.getItem(TOKEN_KEY);
  } catch {
    return null;
  }
}

export function setStoredToken(token: string): void {
  sessionStorage.setItem(TOKEN_KEY, token);
}

export function clearStoredToken(): void {
  sessionStorage.removeItem(TOKEN_KEY);
}

export async function fetchSession(token: string): Promise<{ ok: boolean; email?: string }> {
  try {
    const response = await fetch("/api/auth-session", {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) return { ok: false };
    const data = (await response.json()) as { email?: string };
    if (typeof data.email !== "string" || !data.email) return { ok: false };
    return { ok: true, email: data.email };
  } catch {
    return { ok: false };
  }
}

export async function requestOtp(email: string): Promise<{ ok: true } | { ok: false; status: number }> {
  const response = await fetch("/api/auth-request-otp", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ email: email.trim() }),
  });
  if (response.ok) return { ok: true };
  return { ok: false, status: response.status };
}

export async function verifyOtp(
  email: string,
  code: string,
): Promise<{ ok: true; token: string } | { ok: false; status: number }> {
  const response = await fetch("/api/auth-verify-otp", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ email: email.trim(), code }),
  });
  if (!response.ok) return { ok: false, status: response.status };
  const data = (await response.json()) as { token?: string };
  if (!data.token) return { ok: false, status: 500 };
  return { ok: true, token: data.token };
}
