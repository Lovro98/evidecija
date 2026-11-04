// AI asistent u aplikaciji: prosljeđuje razgovor (i priložene dokumente) Claude API-ju.
// Smije samo admin. API ključ je u tajni ANTHROPIC_API_KEY (Supabase → Edge Functions → Secrets),
// nikad u aplikaciji. Svaki poziv se zapisuje u ai_usage; mjesečni limit u AI_MONTHLY_LIMIT_USD.
// (Kopija onoga što je objavljeno u Supabase kao Edge funkcija "ai-assistant".)
import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "npm:@supabase/supabase-js@2";

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { ...cors, "Content-Type": "application/json" } });

const MODEL = Deno.env.get("AI_MODEL") || "claude-sonnet-5-5";
const PRICE_IN = Number(Deno.env.get("AI_PRICE_IN_PER_M") || 2);   // USD / milijun ulaznih tokena
const PRICE_OUT = Number(Deno.env.get("AI_PRICE_OUT_PER_M") || 10); // USD / milijun izlaznih tokena
const LIMIT_USD = Number(Deno.env.get("AI_MONTHLY_LIMIT_USD") || 30);

const SYSTEM = `Ti si poslovni asistent u aplikaciji "Evidencija rada" agencije koja zapošljava radnike u hotelima u Češkoj i Hrvatskoj.
Korisnik je admin firme. Pomažeš čitati i provjeravati dokumente (tablice sati, fakture, ugovore, platne liste), uspoređivati ih s podacima iz aplikacije, računati i pisati poruke.
Pravila:
- Odgovaraj na jeziku na kojem ti korisnik piše (hrvatski, češki, ukrajinski ili engleski). Poruke za radnike/hotele piši na jeziku primatelja kad ga korisnik navede.
- Budi kratak i konkretan; brojeve i razlike prikaži jasno (tablica ili popis kad pomaže).
- Ako ti je priložen sažetak podataka iz aplikacije, koristi ga, ali reci kad nešto u njemu ne piše — nemoj izmišljati brojeve.
- Ne možeš sam slati poruke, mijenjati podatke ni potvrđivati radnike/cijene; možeš samo predložiti tekst ili korake.
- Za pravna i porezna pitanja daj informaciju, ali napomeni da konačnu odluku treba provjeriti s knjigovođom ili pravnikom.`;

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors });
  if (req.method !== "POST") return json({ error: "Samo POST" }, 405);
  try {
    const url = Deno.env.get("SUPABASE_URL")!;
    const admin = createClient(url, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
    const token = (req.headers.get("Authorization") || "").replace(/^Bearer\s+/i, "");
    const { data: u, error: ue } = await admin.auth.getUser(token);
    if (ue || !u?.user) return json({ error: "Nisi prijavljen." }, 401);
    const { data: me } = await admin.from("profiles").select("role, name").eq("id", u.user.id).single();
    if (me?.role !== "admin") return json({ error: "AI asistent je zasad samo za admine." }, 403);

    const key = Deno.env.get("ANTHROPIC_API_KEY");
    if (!key) return json({ error: "AI asistent još nije uključen: u Supabaseu treba dodati tajnu ANTHROPIC_API_KEY.", code: "no_key" }, 503);

    // mjesečna potrošnja (cijela firma)
    const start = new Date(); start.setUTCDate(1); start.setUTCHours(0, 0, 0, 0);
    const { data: used } = await admin.from("ai_usage").select("cost_usd").gte("created_at", start.toISOString());
    const spent = (used || []).reduce((s: number, r: { cost_usd: number }) => s + Number(r.cost_usd || 0), 0);
    if (spent >= LIMIT_USD) return json({ error: `Dosegnut je mjesečni limit AI asistenta (${LIMIT_USD} $). Limit se mijenja u tajni AI_MONTHLY_LIMIT_USD.`, code: "limit", spent }, 429);

    const body = await req.json();
    const messages = Array.isArray(body.messages) ? body.messages : [];
    if (!messages.length) return json({ error: "Prazna poruka." }, 400);
    const context = typeof body.context === "string" && body.context.trim()
      ? `\n\nSažetak podataka iz aplikacije (stanje ${new Date().toISOString().slice(0, 10)}):\n${body.context.slice(0, 400000)}`
      : "";

    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json" },
      body: JSON.stringify({ model: MODEL, max_tokens: 4096, system: SYSTEM + context, messages }),
    });
    const out = await r.json();
    if (!r.ok) {
      const msg = out?.error?.message || `Greška API-ja (${r.status})`;
      return json({ error: msg }, r.status === 429 ? 429 : 502);
    }
    const inTok = Number(out?.usage?.input_tokens || 0) + Number(out?.usage?.cache_read_input_tokens || 0) + Number(out?.usage?.cache_creation_input_tokens || 0);
    const outTok = Number(out?.usage?.output_tokens || 0);
    const cost = (inTok * PRICE_IN + outTok * PRICE_OUT) / 1e6;
    await admin.from("ai_usage").insert({ user_id: u.user.id, model: out?.model || MODEL, input_tokens: inTok, output_tokens: outTok, cost_usd: cost });
    const text = (out?.content || []).filter((c: { type: string }) => c.type === "text").map((c: { text: string }) => c.text).join("\n");
    return json({ text, usage: { input: inTok, output: outTok, cost }, spent: spent + cost, limit: LIMIT_USD });
  } catch (e) {
    return json({ error: String((e as Error).message || e) }, 500);
  }
});
