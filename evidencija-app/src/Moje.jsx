import { useState, useEffect, useRef } from "react";
import { TR_CS } from "./i18n_cs.js";
import { TR_UK } from "./i18n_uk.js";

// Radnički portal (/moje): radnik se prijavi s ID + PIN i vidi svoje isplaćene obračune.
// Ne učitava aplikaciju i nema pristup bazi — sve ide preko edge funkcije "worker-portal".

const FN_URL = `${import.meta.env.VITE_SUPABASE_URL}/functions/v1/worker-portal`;
const ANON = import.meta.env.VITE_SUPABASE_ANON_KEY;
const IDLE_MS = 10 * 60 * 1000; // nakon 10 min bez dodira odjava

const LANGS = [["HR", "Hrvatski"], ["CZ", "Čeština"], ["UK", "Українська"]];
const MONTHS = {
  HR: ["Siječanj", "Veljača", "Ožujak", "Travanj", "Svibanj", "Lipanj", "Srpanj", "Kolovoz", "Rujan", "Listopad", "Studeni", "Prosinac"],
  CZ: ["Leden", "Únor", "Březen", "Duben", "Květen", "Červen", "Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec"],
  UK: ["Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень", "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень"],
};
const T = {
  "Moje plaće": ["Moje výplaty", "Мої зарплати"],
  "Prijava": ["Přihlášení", "Вхід"],
  "Tvoj ID": ["Vaše ID", "Ваш ID"],
  "Prijavi se": ["Přihlásit se", "Увійти"],
  "Krivi ID ili PIN.": ["Špatné ID nebo PIN.", "Неправильний ID або PIN."],
  "Previše pokušaja. Pokušaj ponovno kasnije.": ["Příliš mnoho pokusů. Zkuste to později.", "Забагато спроб. Спробуйте пізніше."],
  "Greška veze. Pokušaj ponovno.": ["Chyba spojení. Zkuste to znovu.", "Помилка з'єднання. Спробуйте ще раз."],
  "ID i PIN dobiješ od poslodavca.": ["ID a PIN dostanete od zaměstnavatele.", "ID і PIN ви отримаєте від роботодавця."],
  "Odjava": ["Odhlásit", "Вийти"],
  "Isplaćeno": ["Vyplaceno", "Виплачено"],
  "Još nema isplaćenih mjeseci.": ["Zatím žádné vyplacené měsíce.", "Ще немає виплачених місяців."],
  "📄 Potpisana potvrda (PDF)": ["📄 Podepsané potvrzení (PDF)", "📄 Підписане підтвердження (PDF)"],
  "Detaljni obračun za ovaj mjesec nije dostupan.": ["Podrobné vyúčtování za tento měsíc není k dispozici.", "Детальний розрахунок за цей місяць недоступний."],
  "‹ Natrag": ["‹ Zpět", "‹ Назад"],
  "Prijavljujem…": ["Přihlašuji…", "Вхід…"],
  "Ovdje vidiš svoje isplaćene plaće po mjesecima.": ["Zde vidíte své vyplacené výplaty po měsících.", "Тут ви бачите свої виплачені зарплати за місяцями."],
};
const DICT = { CZ: TR_CS, UK: TR_UK };
const initialLang = () => {
  try { const s = localStorage.getItem("moje_lang"); if (s && MONTHS[s]) return s; } catch {}
  const n = (navigator.language || "").toLowerCase();
  return n.startsWith("uk") || n.startsWith("ru") ? "UK" : n.startsWith("cs") || n.startsWith("sk") ? "CZ" : "HR";
};
const store = (k, v) => { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch {} };
const read = (k) => { try { return localStorage.getItem(k) || ""; } catch { return ""; } };

const C = { bg: "#F4F5F1", card: "#fff", ink: "#1C2521", sub: "#5C6862", line: "#E3E6E0", green: "#0E6B4F", greenSoft: "#E4F1EB", red: "#B3261E", redSoft: "#FBE9E7" };
const money = (v, cur) => (cur === "CZK"
  ? `${Math.round(v).toLocaleString("cs-CZ")} Kč`
  : `${Number(v).toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} €`);

export default function Moje() {
  const [lang, setLangState] = useState(initialLang);
  const [code, setCode] = useState(() => { // ID iz QR koda ima prednost
    try { const q = new URLSearchParams(window.location.search).get("id"); if (q) return q.trim().toUpperCase().slice(0, 12); } catch {}
    return read("moje_code");
  });
  const pinRef = useRef(null);
  useEffect(() => { if (code && pinRef.current) pinRef.current.focus(); }, []);
  const [pin, setPin] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [me, setMe] = useState(null); // { name, position, months: [...] }
  const [open, setOpen] = useState(null); // month
  const idle = useRef(null);

  const t = (s) => { const i = lang === "CZ" ? 0 : lang === "UK" ? 1 : -1; return i < 0 ? s : (T[s]?.[i] || s); };
  const setLang = (l) => { setLangState(l); store("moje_lang", l); };
  const monthLabel = (mo) => { const [y, m] = mo.split("-").map(Number); return `${MONTHS[lang][m - 1]} ${y}.`; };
  useEffect(() => { document.documentElement.lang = lang === "UK" ? "uk" : lang === "CZ" ? "cs" : "hr"; document.title = t("Moje plaće"); }, [lang]);

  const logout = () => { setMe(null); setOpen(null); setPin(""); };
  // automatska odjava nakon neaktivnosti (dijeljeni mobitel / tablet)
  useEffect(() => {
    if (!me) return;
    const reset = () => { clearTimeout(idle.current); idle.current = setTimeout(logout, IDLE_MS); };
    reset();
    const ev = ["pointerdown", "keydown", "scroll"];
    ev.forEach((e) => window.addEventListener(e, reset, { passive: true }));
    return () => { clearTimeout(idle.current); ev.forEach((e) => window.removeEventListener(e, reset)); };
  }, [me]);

  const login = async (e) => {
    e.preventDefault();
    if (busy) return;
    setErr(""); setBusy(true);
    try {
      const res = await fetch(FN_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json", apikey: ANON, Authorization: `Bearer ${ANON}` },
        body: JSON.stringify({ code: code.trim().toUpperCase(), pin: pin.trim() }),
      });
      const body = await res.json().catch(() => ({}));
      if (res.status === 429) setErr(t("Previše pokušaja. Pokušaj ponovno kasnije."));
      else if (res.status === 401 || res.status === 400) setErr(t("Krivi ID ili PIN."));
      else if (!res.ok) setErr(t("Greška veze. Pokušaj ponovno."));
      else { setMe(body); setPin(""); store("moje_code", code.trim().toUpperCase()); }
    } catch { setErr(t("Greška veze. Pokušaj ponovno.")); }
    setBusy(false);
  };

  // oznake u obračunu su ⟦hrvatski tekst⟧, ⟦@YYYY-MM⟧ je razdoblje
  const translateSlip = (html) => html
    .replace(/⟦@(\d{4}-\d{2})⟧/g, (_, mo) => monthLabel(mo))
    .replace(/⟦([^⟧]*)⟧/g, (_, k) => (lang === "HR" ? k : DICT[lang]?.[k] || k));

  const langBar = (
    <div style={{ display: "flex", gap: 6, justifyContent: "center", flexWrap: "wrap" }}>
      {LANGS.map(([k, label]) => (
        <button key={k} onClick={() => setLang(k)} style={{ border: `1.5px solid ${lang === k ? C.green : C.line}`, background: lang === k ? C.greenSoft : "#fff",
          color: lang === k ? C.green : C.sub, borderRadius: 999, padding: "7px 13px", fontWeight: 700, fontSize: 14, cursor: "pointer" }}>{label}</button>
      ))}
    </div>
  );

  const page = { minHeight: "100vh", background: C.bg, color: C.ink, fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, sans-serif",
    padding: "calc(16px + env(safe-area-inset-top)) 16px calc(24px + env(safe-area-inset-bottom))", boxSizing: "border-box" };
  const wrap = { maxWidth: 720, margin: "0 auto", display: "flex", flexDirection: "column", gap: 14 };
  const card = { background: C.card, border: `1px solid ${C.line}`, borderRadius: 16, padding: 18 };
  const input = { width: "100%", boxSizing: "border-box", fontSize: 20, padding: "12px 14px", border: `1.5px solid ${C.line}`, borderRadius: 12, letterSpacing: 2, fontWeight: 700 };
  const btn = { width: "100%", minHeight: 52, fontSize: 17, fontWeight: 800, border: "none", borderRadius: 12, background: C.green, color: "#fff", cursor: "pointer" };

  if (!me) return (
    <div style={page}>
      <div style={{ ...wrap, maxWidth: 420 }}>
        <div style={{ textAlign: "center", marginTop: 20 }}>
          <div style={{ fontSize: 40 }}>💶</div>
          <h1 style={{ margin: "6px 0 4px", fontSize: 26 }}>{t("Moje plaće")}</h1>
          <div style={{ color: C.sub, fontSize: 15 }}>{t("Ovdje vidiš svoje isplaćene plaće po mjesecima.")}</div>
        </div>
        {langBar}
        <form onSubmit={login} style={{ ...card, display: "flex", flexDirection: "column", gap: 12 }}>
          <label style={{ fontWeight: 700, fontSize: 14, color: C.sub }}>{t("Tvoj ID")}
            <input value={code} onChange={(e) => setCode(e.target.value.toUpperCase())} autoCapitalize="characters" autoComplete="username"
              autoCorrect="off" spellCheck={false} placeholder="ABC-123" style={{ ...input, marginTop: 6 }} />
          </label>
          <label style={{ fontWeight: 700, fontSize: 14, color: C.sub }}>PIN
            <input value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, "").slice(0, 6))} type="password" inputMode="numeric"
              ref={pinRef} autoComplete="current-password" placeholder="••••••" style={{ ...input, marginTop: 6 }} />
          </label>
          {err && <div style={{ background: C.redSoft, color: C.red, borderRadius: 10, padding: "10px 12px", fontWeight: 600 }}>{err}</div>}
          <button type="submit" disabled={busy || !code.trim() || pin.length < 4} style={{ ...btn, opacity: busy || !code.trim() || pin.length < 4 ? 0.55 : 1 }}>
            {busy ? t("Prijavljujem…") : t("Prijavi se")}
          </button>
          <div style={{ color: C.sub, fontSize: 13.5, textAlign: "center" }}>{t("ID i PIN dobiješ od poslodavca.")}</div>
        </form>
      </div>
    </div>
  );

  const cur = open ? me.months.find((x) => x.month === open) : null;
  const amountText = (x) => x.net || [x.amount ? money(x.amount, "EUR") : "", x.amountCzk ? money(x.amountCzk, "CZK") : ""].filter(Boolean).join(" + ") || money(0, "EUR");

  return (
    <div style={page}>
      <div style={wrap}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          {cur && <button onClick={() => setOpen(null)} style={{ border: "none", background: "none", color: C.green, fontWeight: 800, fontSize: 16, cursor: "pointer", padding: "6px 0" }}>{t("‹ Natrag")}</button>}
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontWeight: 800, fontSize: 19, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{me.name}</div>
            {me.position && <div style={{ color: C.sub, fontSize: 13.5 }}>{me.position}</div>}
          </div>
          <button onClick={logout} style={{ border: `1.5px solid ${C.line}`, background: "#fff", borderRadius: 10, padding: "8px 12px", fontWeight: 700, color: C.sub, cursor: "pointer" }}>{t("Odjava")}</button>
        </div>
        {langBar}

        {!cur && (me.months.length === 0
          ? <div style={{ ...card, textAlign: "center", color: C.sub }}>{t("Još nema isplaćenih mjeseci.")}</div>
          : me.months.map((x) => (
            <button key={x.month} onClick={() => setOpen(x.month)} style={{ ...card, display: "flex", alignItems: "center", gap: 12, textAlign: "left", cursor: "pointer", font: "inherit", color: C.ink }}>
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 800, fontSize: 17 }}>{monthLabel(x.month)}</div>
                <div style={{ color: C.green, fontWeight: 700, fontSize: 13.5, marginTop: 2 }}>✓ {t("Isplaćeno")}</div>
              </div>
              <div style={{ fontWeight: 800, fontSize: 18, whiteSpace: "nowrap" }}>{amountText(x)}</div>
              <div style={{ color: C.sub, fontSize: 22 }}>›</div>
            </button>
          )))}

        {cur && (
          <>
            <div style={{ ...card, display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 800, fontSize: 18 }}>{monthLabel(cur.month)}</div>
                <div style={{ color: C.green, fontWeight: 700, fontSize: 13.5 }}>✓ {t("Isplaćeno")}</div>
              </div>
              <div style={{ fontWeight: 800, fontSize: 22 }}>{amountText(cur)}</div>
            </div>
            {cur.pdf && <a href={cur.pdf} target="_blank" rel="noopener noreferrer" style={{ ...btn, display: "flex", alignItems: "center", justifyContent: "center", textDecoration: "none", background: "#fff", color: C.green, border: `1.5px solid ${C.green}` }}>{t("📄 Potpisana potvrda (PDF)")}</a>}
            {cur.html
              ? <Slip html={translateSlip(cur.html)} />
              : <div style={{ ...card, color: C.sub, textAlign: "center" }}>{t("Detaljni obračun za ovaj mjesec nije dostupan.")}</div>}
          </>
        )}
      </div>
    </div>
  );
}

// obračun u iframeu bez skripti (sandbox), visina se prilagodi sadržaju
function Slip({ html }) {
  const ref = useRef(null);
  const [h, setH] = useState(600);
  const fit = () => { try { const d = ref.current?.contentDocument; if (d?.body) setH(d.documentElement.scrollHeight + 4); } catch {} };
  useEffect(() => { const id = setInterval(fit, 500); return () => clearInterval(id); }, [html]);
  return (
    <div style={{ background: "#fff", border: "1px solid #E3E6E0", borderRadius: 16, overflow: "hidden" }}>
      <iframe ref={ref} title="obracun" sandbox="allow-same-origin" srcDoc={html} onLoad={fit}
        style={{ width: "100%", height: h, border: "none", display: "block" }} />
    </div>
  );
}
