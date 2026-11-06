// Radnički portal (/moje): radnik pošalje ID + PIN, dobije SAMO svoje isplaćene (odobrene) obračune.
// Radnici nemaju Supabase račun — provjera ide preko portal_login (bcrypt + zaključavanje nakon krivih PIN-ova).
// (Kopija onoga što je objavljeno u Supabase kao Edge funkcija "worker-portal", verify_jwt = false.)
import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "npm:@supabase/supabase-js@2";

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { ...cors, "Content-Type": "application/json", "Cache-Control": "no-store" } });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors });
  if (req.method !== "POST") return json({ error: "method" }, 405);
  try {
    const body = await req.json().catch(() => ({}));
    const code = typeof body.code === "string" ? body.code.slice(0, 20) : "";
    const pin = typeof body.pin === "string" ? body.pin.slice(0, 12) : "";
    if (!code || !pin) return json({ error: "bad" }, 400);

    const sb = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!, { auth: { persistSession: false } });
    const { data: lg, error: lgErr } = await sb.rpc("portal_login", { p_code: code, p_pin: pin });
    if (lgErr) throw lgErr;
    if (!lg?.ok) {
      await new Promise((r) => setTimeout(r, 500)); // usporava pogađanje
      return json({ error: lg?.reason || "bad", until: lg?.until || null }, lg?.reason === "locked" ? 429 : 401);
    }
    const wid = lg.worker_id as string;

    const [w, pays, snaps, docs] = await Promise.all([
      sb.from("workers").select("name, position").eq("id", wid).single(),
      sb.from("payouts").select("month, amount, amount_czk, approved_at").eq("worker_id", wid).eq("approved", true).is("deleted_at", null),
      sb.from("payslip_snapshots").select("month, html, net").eq("worker_id", wid),
      sb.from("worker_documents").select("file_path, note, created_at").eq("worker_id", wid).eq("doc_type", "isplata").is("deleted_at", null),
    ]);
    for (const r of [w, pays, snaps, docs]) if (r.error) throw r.error;

    const months = [...new Map((pays.data || []).map((p) => [p.month, p])).values()]
      .sort((a, b) => b.month.localeCompare(a.month)).slice(0, 24);
    const out = [];
    for (const p of months) {
      const s = (snaps.data || []).find((x) => x.month === p.month);
      const d = (docs.data || []).filter((x) => x.file_path && (x.note || "").startsWith(`Isplata ${p.month}`))
        .sort((a, b) => (b.created_at || "").localeCompare(a.created_at || ""))[0];
      let pdf = null;
      if (d) {
        const { data: su } = await sb.storage.from("docs").createSignedUrl(d.file_path, 3600);
        pdf = su?.signedUrl || null;
      }
      out.push({ month: p.month, amount: Number(p.amount) || 0, amountCzk: Number(p.amount_czk) || 0, paidAt: p.approved_at,
        net: s?.net || null, html: s?.html || null, pdf });
    }
    return json({ name: w.data?.name || "", position: w.data?.position || "", months: out });
  } catch (e) {
    console.error(e);
    return json({ error: "server" }, 500);
  }
});
