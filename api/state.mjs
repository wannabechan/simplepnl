import { neon } from "@neondatabase/serverless";

const APP_STATE_ID = "simplepnl-main";

async function ensureTable(sql) {
  await sql`
    create table if not exists app_state (
      id text primary key,
      data jsonb not null,
      updated_at timestamptz not null default now()
    )
  `;
}

export default async function handler(req, res) {
  try {
    const databaseUrl = typeof process.env.DATABASE_URL === "string" ? process.env.DATABASE_URL.trim() : "";
    if (!databaseUrl) {
      res.status(500).json({ error: "DATABASE_URL is not configured." });
      return;
    }

    const sql = neon(databaseUrl);
    await ensureTable(sql);

    if (req.method === "GET") {
      const rows = await sql`select data from app_state where id = ${APP_STATE_ID} limit 1`;
      res.status(200).json({ data: rows[0]?.data ?? [] });
      return;
    }

    if (req.method === "PUT") {
      let rawBody;
      try {
        rawBody = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
      } catch {
        res.status(400).json({ error: "Invalid JSON body" });
        return;
      }
      const payload = Array.isArray(rawBody?.data) ? rawBody.data : [];
      await sql`
        insert into app_state (id, data, updated_at)
        values (${APP_STATE_ID}, ${JSON.stringify(payload)}::jsonb, now())
        on conflict (id)
        do update set data = excluded.data, updated_at = now()
      `;
      res.status(200).json({ ok: true });
      return;
    }

    res.setHeader("Allow", "GET, PUT");
    res.status(405).json({ error: "Method not allowed" });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[api/state]", message, err);
    if (!res.headersSent) {
      res.status(500).json({ error: message });
    }
  }
}
