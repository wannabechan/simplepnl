import { neon } from "@neondatabase/serverless";

/** @param {unknown} s */
export function normalizeEmail(s) {
  return String(s ?? "").trim().toLowerCase();
}

export function getAccountAdmin() {
  return normalizeEmail(process.env.ACCOUNT_ADMIN ?? "");
}

export function getSqlOrNull() {
  const databaseUrl = typeof process.env.DATABASE_URL === "string" ? process.env.DATABASE_URL.trim() : "";
  if (!databaseUrl) return null;
  return neon(databaseUrl);
}

/** @param {ReturnType<typeof neon>} sql */
export async function ensureAuthTables(sql) {
  await sql`
    create table if not exists login_otp (
      email text primary key,
      code text not null,
      expires_at timestamptz not null
    )
  `;
  await sql`
    create table if not exists login_sessions (
      token text primary key,
      email text not null,
      expires_at timestamptz not null
    )
  `;
}
