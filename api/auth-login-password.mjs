import crypto from "node:crypto";
import { getAccountAdmin, getSqlOrNull, ensureAuthTables, normalizeEmail } from "./_authDb.mjs";

function normalizeEmailInput(raw) {
  return String(raw ?? "").trim().toLowerCase();
}

function getExpectedPassword() {
  return typeof process.env.ACCOUNT_ADMIN_PW === "string" ? process.env.ACCOUNT_ADMIN_PW : "";
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    const admin = getAccountAdmin();
    const expectedPw = getExpectedPassword();
    if (!admin) {
      res.status(500).json({ error: "ACCOUNT_ADMIN is not configured." });
      return;
    }
    if (!expectedPw.trim()) {
      res.status(500).json({ error: "ACCOUNT_ADMIN_PW is not configured." });
      return;
    }

    let body;
    try {
      body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    } catch {
      res.status(400).json({ error: "Invalid JSON body" });
      return;
    }

    const email = normalizeEmailInput(body?.email);
    const password = typeof body?.password === "string" ? body.password : "";

    if (!email || !email.includes("@")) {
      res.status(400).json({ error: "Invalid email" });
      return;
    }

    if (email !== admin) {
      res.status(403).json({ error: "forbidden" });
      return;
    }

    if (password !== expectedPw) {
      res.status(401).json({ error: "invalid_password" });
      return;
    }

    const sql = getSqlOrNull();
    if (!sql) {
      res.status(500).json({ error: "DATABASE_URL is not configured." });
      return;
    }

    await ensureAuthTables(sql);

    const token = crypto.randomBytes(32).toString("hex");
    const sessionExpires = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);

    await sql`
      insert into login_sessions (token, email, expires_at)
      values (${token}, ${normalizeEmail(email)}, ${sessionExpires.toISOString()})
    `;

    res.status(200).json({ ok: true, token });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[auth-login-password]", message, err);
    if (!res.headersSent) {
      res.status(500).json({ error: message });
    }
  }
}
