import crypto from "node:crypto";
import { getSqlOrNull, ensureAuthTables, normalizeEmail } from "./_authDb.mjs";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    let body;
    try {
      body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    } catch {
      res.status(400).json({ error: "Invalid JSON body" });
      return;
    }

    const email = normalizeEmail(body?.email);
    const code = String(body?.code ?? "").replace(/\D/g, "").slice(0, 6);

    if (!email || !email.includes("@") || code.length !== 6) {
      res.status(400).json({ error: "Invalid email or code" });
      return;
    }

    const sql = getSqlOrNull();
    if (!sql) {
      res.status(500).json({ error: "DATABASE_URL is not configured." });
      return;
    }

    await ensureAuthTables(sql);

    const rows = await sql`
      select code, expires_at from login_otp where email = ${email} limit 1
    `;
    const row = rows[0];
    if (!row) {
      res.status(401).json({ error: "No pending code for this email." });
      return;
    }

    const expires = row.expires_at ? new Date(row.expires_at) : null;
    if (!expires || Number.isNaN(expires.getTime()) || expires.getTime() < Date.now()) {
      await sql`delete from login_otp where email = ${email}`;
      res.status(401).json({ error: "Code expired." });
      return;
    }

    if (String(row.code) !== code) {
      res.status(401).json({ error: "Invalid code." });
      return;
    }

    await sql`delete from login_otp where email = ${email}`;

    const token = crypto.randomBytes(32).toString("hex");
    const sessionExpires = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);

    await sql`
      insert into login_sessions (token, email, expires_at)
      values (${token}, ${email}, ${sessionExpires.toISOString()})
    `;

    res.status(200).json({ ok: true, token });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[auth-verify-otp]", message, err);
    if (!res.headersSent) {
      res.status(500).json({ error: message });
    }
  }
}
