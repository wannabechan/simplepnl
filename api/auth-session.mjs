import { getSqlOrNull, ensureAuthTables, normalizeEmail } from "./_authDb.mjs";

function extractBearer(req) {
  const raw = req.headers?.authorization ?? req.headers?.Authorization;
  if (typeof raw !== "string") return "";
  const m = raw.match(/^Bearer\s+(.+)$/i);
  return m ? m[1].trim() : "";
}

export default async function handler(req, res) {
  if (req.method !== "GET") {
    res.setHeader("Allow", "GET");
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    const token = extractBearer(req);
    if (!token) {
      res.status(401).json({ error: "Missing token" });
      return;
    }

    const sql = getSqlOrNull();
    if (!sql) {
      res.status(500).json({ error: "DATABASE_URL is not configured." });
      return;
    }

    await ensureAuthTables(sql);

    await sql`delete from login_sessions where expires_at < now()`;

    const rows = await sql`
      select email from login_sessions where token = ${token} and expires_at > now() limit 1
    `;
    const email = rows[0]?.email;
    if (!email || typeof email !== "string") {
      res.status(401).json({ error: "Invalid or expired session" });
      return;
    }

    res.status(200).json({ ok: true, email: normalizeEmail(email) });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[auth-session]", message, err);
    if (!res.headersSent) {
      res.status(500).json({ error: message });
    }
  }
}
