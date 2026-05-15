import { getAccountAdmin } from "./_authDb.mjs";

function normalizeEmailInput(raw) {
  return String(raw ?? "").trim().toLowerCase();
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

    let body;
    try {
      body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    } catch {
      res.status(400).json({ error: "Invalid JSON body" });
      return;
    }

    const email = normalizeEmailInput(body?.email);
    if (!email || !email.includes("@")) {
      res.status(400).json({ error: "Invalid email" });
      return;
    }

    if (email !== admin) {
      res.status(403).json({ error: "forbidden" });
      return;
    }

    res.status(200).json({ ok: true });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[auth-check-account]", message, err);
    if (!res.headersSent) {
      res.status(500).json({ error: message });
    }
  }
}
