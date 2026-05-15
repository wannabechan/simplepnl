import { getAccountAdmin, getSqlOrNull, ensureAuthTables, normalizeEmail } from "./_authDb.mjs";

function randomSixDigitCode() {
  return String(Math.floor(Math.random() * 1_000_000)).padStart(6, "0");
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    const admin = getAccountAdmin();
    if (!admin) {
      res.status(500).json({ error: "ACCOUNT_ADMIN is not configured." });
      return;
    }

    const resendKey = typeof process.env.RESEND_API_KEY === "string" ? process.env.RESEND_API_KEY.trim() : "";
    if (!resendKey) {
      res.status(500).json({ error: "RESEND_API_KEY is not configured." });
      return;
    }

    let body;
    try {
      body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    } catch {
      res.status(400).json({ error: "Invalid JSON body" });
      return;
    }

    const email = normalizeEmail(body?.email);
    if (!email || !email.includes("@")) {
      res.status(400).json({ error: "Invalid email" });
      return;
    }

    if (email !== admin) {
      res.status(403).json({ error: "forbidden" });
      return;
    }

    const sql = getSqlOrNull();
    if (!sql) {
      res.status(500).json({ error: "DATABASE_URL is not configured." });
      return;
    }

    await ensureAuthTables(sql);

    const code = randomSixDigitCode();
    const expiresAt = new Date(Date.now() + 10 * 60 * 1000);

    await sql`
      insert into login_otp (email, code, expires_at)
      values (${email}, ${code}, ${expiresAt.toISOString()})
      on conflict (email)
      do update set code = excluded.code, expires_at = excluded.expires_at
    `;

    const fromRaw =
      typeof process.env.RESEND_FROM_EMAIL === "string" && process.env.RESEND_FROM_EMAIL.trim()
        ? process.env.RESEND_FROM_EMAIL.trim()
        : "SimplePNL <onboarding@resend.dev>";

    const sendRes = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${resendKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from: fromRaw,
        to: [email],
        subject: "SimplePNL 로그인 인증 코드",
        text: `인증 코드: ${code}\n\n10분 이내에 입력해 주세요.`,
      }),
    });

    if (!sendRes.ok) {
      const errText = await sendRes.text();
      console.error("[auth-request-otp] Resend error", sendRes.status, errText);
      res.status(502).json({ error: "Failed to send email." });
      return;
    }

    res.status(200).json({ ok: true });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[auth-request-otp]", message, err);
    if (!res.headersSent) {
      res.status(500).json({ error: message });
    }
  }
}
