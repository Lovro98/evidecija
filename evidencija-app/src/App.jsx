import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase.js";

/* ================================================================== */
/*  EVIDENCIJA RADA — puna verzija                                     */
/* ================================================================== */

const S = {
  bg: "#F4F5F1", card: "#FFFFFF", ink: "#1C2521", sub: "#5C6862", line: "#E1E5DF",
  green: "#0E6B4F", greenSoft: "#E3F0EA", amber: "#B97F1E", amberSoft: "#FBF1DC",
  red: "#B3402F", redSoft: "#F9E8E4", blue: "#2B5D8A", blueSoft: "#E6EEF5",
};

const eur = (n) => (isNaN(n) ? 0 : n).toLocaleString("hr-HR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €";
const czk = (n) => (isNaN(n) ? 0 : n).toLocaleString("cs-CZ", { minimumFractionDigits: 0, maximumFractionDigits: 2 }) + " Kč";
const money = (n, c) => (c === "CZK" ? czk(n) : eur(n));
const FLAG = { HR: "🇭🇷", CZ: "🇨🇿" };
const wCur = (w) => (w?.rateCurrency === "CZK" ? "CZK" : "EUR");
const countryCur = (country) => (country === "CZ" ? "CZK" : "EUR");
const byName = (a, b) => a.name.localeCompare(b.name, "hr");
const sortedWorkers = (list) => [...list].sort(byName);
const sortedObjects = (list) => [...list].sort(byName);
const workersAtObject = (data, objectId) => {
  const active = data.workers.filter((w) => !w.archived);
  if (!objectId) return sortedWorkers(active);
  const withLogs = new Set(data.logs.filter((l) => l.objectId === objectId).map((l) => l.workerId));
  return sortedWorkers(active.filter((w) => w.objectId === objectId || withLogs.has(w.id)));
};
const CurChips = ({ value, onChange, small }) => (
  <div style={{ display: "flex", gap: 6 }}>
    {[["EUR", "€"], ["CZK", "Kč"]].map(([id, label]) => (
      <button key={id} type="button" onClick={() => onChange(id)} style={{
        padding: small ? "5px 10px" : "7px 14px", borderRadius: 999, fontWeight: 700, fontSize: small ? 12.5 : 13, cursor: "pointer",
        background: value === id ? "#0E6B4F" : "#fff", color: value === id ? "#fff" : "#5C6862",
        border: `1px solid ${value === id ? "#0E6B4F" : "#E1E5DF"}` }}>{label}</button>
    ))}
  </div>
);
const COUNTRY_NAME = { HR: "Hrvatska", CZ: "Češka" };
const round2 = (n) => Math.round(n * 100) / 100;
const MONTHS = ["Siječanj","Veljača","Ožujak","Travanj","Svibanj","Lipanj","Srpanj","Kolovoz","Rujan","Listopad","Studeni","Prosinac"];
const MONTHS_SHORT = ["sij","vlj","ožu","tra","svi","lip","srp","kol","ruj","lis","stu","pro"];

function hoursBetween(from, to) {
  if (!from || !to) return 0;
  const [fh, fm] = from.split(":").map(Number);
  const [th, tm] = to.split(":").map(Number);
  let mins = th * 60 + tm - (fh * 60 + fm);
  if (mins < 0) mins += 24 * 60;
  return round2(mins / 60);
}
const fmtH = (h) => { const w = Math.floor(h); const m = Math.round((h - w) * 60); return m ? `${w} h ${m} min` : `${w} h`; };
const fmtDate = (iso) => { if (!iso) return ""; const [y, m, d] = iso.split("-"); return `${Number(d)}.${Number(m)}.${y}.`; };
const fmtDT = (iso) => { const d = new Date(iso); return `${d.getDate()}.${d.getMonth() + 1}. ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`; };
const monthKey = (iso) => (iso || "").slice(0, 7);
const todayISO = () => new Date().toISOString().slice(0, 10);
const curMonth = () => { const n = new Date(); return `${n.getFullYear()}-${String(n.getMonth() + 1).padStart(2, "0")}`; };
const logSpan = (l) => (l.from && l.to ? `${l.from}–${l.to}` : l.monthly ? "mjesečni zbroj" : "upis sati");
const parseNum = (v) => parseFloat(String(v || "").replace(",", "."));
const TYPE_LABEL = { avans: "Avans", bonus: "Bonus", gorivo: "Gorivo", ostalo: "Ostali trošak" };

/* ---------- jezici (HR / CZ / UK) ---------- */
const LANGS = ["HR", "CZ", "UK"];
const LANG_FLAG = { HR: "🇭🇷", CZ: "🇨🇿", UK: "🇺🇦" };
const LANG_NAME = { HR: "Hrvatski", CZ: "Čeština", UK: "Українська" };
const DICT = {
  HR: {
    appTitle: "Evidencija rada", appSubtitle: "Radnici · Sati · Obračun", logout: "Odjava",
    tabRadnici: "Radnici", tabObjekti: "Objekti", tabImenik: "Imenik", tabSati: "Sati", tabIsplate: "Isplate", tabObracun: "Obračun",
    loginTitle: "Prijava", email: "E-mail", password: "Lozinka", loginBtn: "Prijavi se", loginBusy: "Prijavljujem…",
    loginErr: "Neispravan e-mail ili lozinka.",
    noAccount: "Nemaš račun? Račune otvara isključivo poslodavac — javi mu se i dobit ćeš e-mail i lozinku.",
    workersTitle: "Radnici", addWorker: "+ Novi radnik", close: "Zatvori",
    fullName: "Ime i prezime *", phone: "Broj telefona", rate: "Satnica", mainObject: "Glavni objekt", note: "Napomena",
    saveWorker: "Spremi radnika", noWorkersYet: "Još nema radnika. Dodaj prvog gumbom + Novi radnik.",
    formerWorkers: "Bivši radnici", active: "Aktivan", withoutRate: "bez satnice", thisMonth: "ovaj mj.",
    objectsTitle: "Objekti", newObjectPh: "Naziv novog objekta u:", addBtn: "Dodaj",
    directoryTitle: "Imenik radnika", searchPh: "🔍 Traži po imenu ili broju…",
    hoursTitle: "Upis radnih sati", object: "Objekt", worker: "Radnik", date: "Datum", from: "Od", to: "Do",
    total: "Ukupno", addHours: "Upiši sate", recent: "Zadnji upisi", choose: "— odaberi —",
    paymentsTitle: "Avansi, bonusi i troškovi", forWorker: "👷 Za radnika", forObject: "🏨 Za objekt",
    type: "Vrsta", amount: "Iznos (€)", save: "Spremi",
    monthBtn: "Mjesec", yearBtn: "Godina", showFor: "Prikaz za", allObjects: "🏨 Svi objekti",
    totalHours: "Ukupno sati", earningsBonus: "Zarada + bonusi", toPay: "Za isplatiti", firmCosts: "Troškovi firme",
    excelBtn: "Excel", envelopesBtn: "Koverte", cancel: "Odustani", edit: "Uredi", delete: "Obriši", back: "Natrag",
  },
  CZ: {
    appTitle: "Evidence práce", appSubtitle: "Pracovníci · Hodiny · Vyúčtování", logout: "Odhlásit se",
    tabRadnici: "Pracovníci", tabObjekti: "Objekty", tabImenik: "Adresář", tabSati: "Hodiny", tabIsplate: "Výplaty", tabObracun: "Vyúčtování",
    loginTitle: "Přihlášení", email: "E-mail", password: "Heslo", loginBtn: "Přihlásit se", loginBusy: "Přihlašuji…",
    loginErr: "Nesprávný e-mail nebo heslo.",
    noAccount: "Nemáte účet? Účty zakládá pouze zaměstnavatel — ozvěte se mu a dostanete e-mail a heslo.",
    workersTitle: "Pracovníci", addWorker: "+ Nový pracovník", close: "Zavřít",
    fullName: "Jméno a příjmení *", phone: "Telefonní číslo", rate: "Sazba", mainObject: "Hlavní objekt", note: "Poznámka",
    saveWorker: "Uložit pracovníka", noWorkersYet: "Zatím žádní pracovníci. Přidejte prvního tlačítkem + Nový pracovník.",
    formerWorkers: "Bývalí pracovníci", active: "Aktivní", withoutRate: "bez sazby", thisMonth: "tento měs.",
    objectsTitle: "Objekty", newObjectPh: "Název nového objektu v:", addBtn: "Přidat",
    directoryTitle: "Adresář pracovníků", searchPh: "🔍 Hledat podle jména nebo čísla…",
    hoursTitle: "Zápis odpracovaných hodin", object: "Objekt", worker: "Pracovník", date: "Datum", from: "Od", to: "Do",
    total: "Celkem", addHours: "Zapsat hodiny", recent: "Poslední zápisy", choose: "— vyberte —",
    paymentsTitle: "Zálohy, bonusy a náklady", forWorker: "👷 Pro pracovníka", forObject: "🏨 Pro objekt",
    type: "Druh", amount: "Částka (€)", save: "Uložit",
    monthBtn: "Měsíc", yearBtn: "Rok", showFor: "Zobrazit za", allObjects: "🏨 Všechny objekty",
    totalHours: "Celkem hodin", earningsBonus: "Výdělek + bonusy", toPay: "K výplatě", firmCosts: "Náklady firmy",
    excelBtn: "Excel", envelopesBtn: "Obálky", cancel: "Zrušit", edit: "Upravit", delete: "Smazat", back: "Zpět",
  },
  UK: {
    appTitle: "Облік роботи", appSubtitle: "Працівники · Години · Розрахунок", logout: "Вийти",
    tabRadnici: "Працівники", tabObjekti: "Об'єкти", tabImenik: "Довідник", tabSati: "Години", tabIsplate: "Виплати", tabObracun: "Розрахунок",
    loginTitle: "Вхід", email: "Е-пошта", password: "Пароль", loginBtn: "Увійти", loginBusy: "Вхід…",
    loginErr: "Невірна електронна пошта або пароль.",
    noAccount: "Немає акаунту? Акаунти створює лише роботодавець — зверніться до нього і отримаєте е-пошту та пароль.",
    workersTitle: "Працівники", addWorker: "+ Новий працівник", close: "Закрити",
    fullName: "Ім'я та прізвище *", phone: "Номер телефону", rate: "Ставка", mainObject: "Основний об'єкт", note: "Примітка",
    saveWorker: "Зберегти працівника", noWorkersYet: "Ще немає працівників. Додайте першого кнопкою + Новий працівник.",
    formerWorkers: "Колишні працівники", active: "Активний", withoutRate: "без ставки", thisMonth: "цього міс.",
    objectsTitle: "Об'єкти", newObjectPh: "Назва нового об'єкта в:", addBtn: "Додати",
    directoryTitle: "Довідник працівників", searchPh: "🔍 Пошук за іменем або номером…",
    hoursTitle: "Облік відпрацьованих годин", object: "Об'єкт", worker: "Працівник", date: "Дата", from: "Від", to: "До",
    total: "Всього", addHours: "Записати години", recent: "Останні записи", choose: "— обрати —",
    paymentsTitle: "Аванси, бонуси та витрати", forWorker: "👷 Для працівника", forObject: "🏨 Для об'єкта",
    type: "Тип", amount: "Сума (€)", save: "Зберегти",
    monthBtn: "Місяць", yearBtn: "Рік", showFor: "Показати для", allObjects: "🏨 Всі об'єкти",
    totalHours: "Всього годин", earningsBonus: "Заробіток + бонуси", toPay: "До виплати", firmCosts: "Витрати фірми",
    excelBtn: "Excel", envelopesBtn: "Конверти", cancel: "Скасувати", edit: "Редагувати", delete: "Видалити", back: "Назад",
  },
};
function useLang() {
  const [lang, setLang] = useState(() => { try { return localStorage.getItem("evidencija_lang") || "HR"; } catch { return "HR"; } });
  useEffect(() => { try { localStorage.setItem("evidencija_lang", lang); } catch {} }, [lang]);
  const t = (k) => (DICT[lang] && DICT[lang][k]) || DICT.HR[k] || k;
  return { lang, setLang, t };
}
function LangSwitch({ lang, setLang, dark }) {
  return (
    <div style={{ display: "flex", gap: 4 }}>
      {LANGS.map((l) => (
        <button key={l} onClick={() => setLang(l)} title={LANG_NAME[l]} style={{
          background: lang === l ? (dark ? "rgba(255,255,255,.22)" : S.green) : (dark ? "rgba(255,255,255,.08)" : "#fff"),
          border: dark ? "1px solid rgba(255,255,255,.25)" : `1px solid ${S.line}`,
          borderRadius: 8, padding: "4px 7px", fontSize: 15, cursor: "pointer", lineHeight: 1 }}>
          {LANG_FLAG[l]}
        </button>
      ))}
    </div>
  );
}
const addDaysISO = (iso, n) => { const d = new Date(iso + "T12:00:00"); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); };

/* ---------- hrvatski praznici i vikendi ---------- */
function easterISO(y) {
  const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4,
    f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30,
    i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7,
    m = Math.floor((a + 11 * h + 22 * l) / 451), mo = Math.floor((h + l - 7 * m + 114) / 31),
    da = ((h + l - 7 * m + 114) % 31) + 1;
  return `${y}-${String(mo).padStart(2, "0")}-${String(da).padStart(2, "0")}`;
}
function holidayName(iso) {
  if (!iso) return null;
  const [y, m, d] = iso.split("-");
  const fixed = {
    "01-01": "Nova godina", "01-06": "Sveta tri kralja", "05-01": "Praznik rada",
    "05-30": "Dan državnosti", "06-22": "Dan antifašističke borbe", "08-05": "Dan pobjede",
    "08-15": "Velika Gospa", "11-01": "Svi sveti", "11-18": "Dan sjećanja",
    "12-25": "Božić", "12-26": "Sveti Stjepan",
  };
  if (fixed[`${m}-${d}`]) return fixed[`${m}-${d}`];
  const e = easterISO(Number(y));
  if (iso === e) return "Uskrs";
  if (iso === addDaysISO(e, 1)) return "Uskrsni ponedjeljak";
  if (iso === addDaysISO(e, 60)) return "Tijelovo";
  return null;
}
const isWeekend = (iso) => { const g = new Date(iso + "T12:00:00").getDay(); return g === 0 || g === 6; };
function DayBadge({ iso }) {
  const h = holidayName(iso);
  if (h) return <Tag color={S.red} bg={S.redSoft}>🎉 Praznik: {h}</Tag>;
  if (isWeekend(iso)) return <Tag color={S.blue} bg={S.blueSoft}>{new Date(iso + "T12:00:00").getDay() === 6 ? "Subota" : "Nedjelja"}</Tag>;
  return null;
}

function rateFor(data, w, date) {
  if (!w) return 0;
  const ch = (data.rateChanges || [])
    .filter((r) => r.workerId === w.id && r.from <= (date || "9999-12-31"))
    .sort((a, b) => a.from.localeCompare(b.from));
  return ch.length ? ch[ch.length - 1].rate : (w.rate || 0);
}
const rateNow = (data, w) => rateFor(data, w, todayISO());
const paidFor = (data, workerId, mo) => (data.payouts || []).find((p) => p.workerId === workerId && p.month === mo);

/* ---------- ispis PDF-a (preglednikov "Spremi kao PDF" / dijeljenje) ---------- */
function printDoc(title, bodyHtml) {
  const w = window.open("", "_blank");
  if (!w) return;
  w.document.write(`<!DOCTYPE html><html lang="hr"><head><meta charset="utf-8"><title>${title}</title><style>
    body{font-family:Arial,Helvetica,sans-serif;color:#1C2521;padding:24px;max-width:720px;margin:0 auto}
    h1{font-size:20px;margin:0 0 2px} h2{font-size:14px;color:#5C6862;font-weight:normal;margin:0 0 14px}
    table{width:100%;border-collapse:collapse;margin:12px 0}
    th,td{border-bottom:1px solid #ddd;padding:6px 8px;text-align:left;font-size:13px}
    th{background:#f2f4f0;font-size:12px} .right{text-align:right}
    .tot td{font-weight:bold;font-size:14px;border-top:2px solid #1C2521}
    .muted{color:#5C6862;font-size:12px} .head{border-bottom:2px solid #1C2521;padding-bottom:10px;margin-bottom:14px}
  </style></head><body>${bodyHtml}<script>setTimeout(function(){window.print()},350)<\/script></body></html>`);
  w.document.close();
}
const firmHead = (st) => `<div class="head"><h1>${st.company_name || "Moja firma"}</h1>
  <div class="muted">${[st.address, st.oib ? "OIB: " + st.oib : "", st.iban ? "IBAN: " + st.iban : ""].filter(Boolean).join(" · ")}</div></div>`;
function printGeneratedInvoice(inv, object, st) {
  const [y, m] = inv.period.split("-").map(Number);
  const periodLabel = `${MONTHS[m - 1]} ${y}.`;
  printDoc(`Faktura ${inv.number}`, `
    ${firmHead(st)}
    <h1>Faktura br. ${inv.number}</h1>
    <h2>${fmtDate(inv.issueDate)} · Kupac: ${object.name}</h2>
    <table><tr><th>Opis</th><th class="right">Sati</th><th class="right">Cijena/h</th><th class="right">Iznos</th></tr>
    <tr><td>Usluge iznajmljivanja radne snage — ${periodLabel}${inv.note ? " · " + inv.note : ""}</td>
      <td class="right">${fmtH(inv.hours)}</td><td class="right">${money(inv.rate, inv.currency)}</td><td class="right">${money(inv.amount, inv.currency)}</td></tr>
    <tr class="tot"><td colspan="3">UKUPNO ZA PLATITI</td><td class="right">${money(inv.amount, inv.currency)}</td></tr></table>
    ${st.iban ? `<div class="muted">Uplata na IBAN: <b>${st.iban}</b></div>` : ""}
    <div class="muted" style="margin-top:20px">Datum dospijeća: ______________________</div>`);
}

/* ================================================================== */
/*  UČITAVANJE PODATAKA                                                */
/* ================================================================== */
async function fetchAll(isAdmin) {
  const live = (q) => q.is("deleted_at", null);
  const [workers, objects, logs, payments, assignments, rateChanges, profiles, payouts] = await Promise.all([
    live(supabase.from("workers").select("*")).order("name"),
    live(supabase.from("objects").select("*")).order("name"),
    live(supabase.from("work_logs").select("*")),
    live(supabase.from("payments").select("*")),
    live(supabase.from("assignments").select("*")),
    live(supabase.from("rate_changes").select("*")),
    supabase.from("profiles").select("*"),
    live(supabase.from("payouts").select("*")),
  ]);
  const err = [workers, objects, logs, payments, assignments, rateChanges, profiles, payouts].find((r) => r.error);
  if (err) throw err.error;

  let billing = [], audit = [], trash = [], members = [], invoicePayments = [], settings = {}, objectInvoices = [];
  if (isAdmin) {
    const [b, a, tw, tl, tp, om, ip, st, oi] = await Promise.all([
      supabase.from("object_billing").select("*"),
      supabase.from("audit_log").select("*").order("at", { ascending: false }).limit(80),
      supabase.from("workers").select("*").not("deleted_at", "is", null),
      supabase.from("work_logs").select("*").not("deleted_at", "is", null).order("deleted_at", { ascending: false }).limit(50),
      supabase.from("payments").select("*").not("deleted_at", "is", null).order("deleted_at", { ascending: false }).limit(50),
      supabase.from("object_members").select("*"),
      live(supabase.from("invoice_payments").select("*")),
      supabase.from("settings").select("*").eq("id", 1).maybeSingle(),
      live(supabase.from("object_invoices").select("*")).order("issue_date", { ascending: false }),
    ]);
    billing = b.data || []; audit = a.data || []; members = om.data || [];
    invoicePayments = ip.data || []; settings = st.data || {}; objectInvoices = oi.data || [];
    trash = [
      ...(tw.data || []).map((r) => ({ table: "workers", row: r, label: `Radnik: ${r.name}` })),
      ...(tl.data || []).map((r) => ({ table: "work_logs", row: r, label: `Sati: ${fmtH(r.hours)} (${fmtDate(r.work_date)})` })),
      ...(tp.data || []).map((r) => ({ table: "payments", row: r, label: `${TYPE_LABEL[r.type] || r.type}: ${eur(r.amount)} (${fmtDate(r.pay_date)})` })),
    ].sort((a, b) => (b.row.deleted_at || "").localeCompare(a.row.deleted_at || ""));
  }
  const billMap = Object.fromEntries(billing.map((b) => [b.object_id, { rate: Number(b.bill_rate) || 0, cur: b.bill_currency === "CZK" ? "CZK" : "EUR" }]));

  return {
    workers: (workers.data || []).map((w) => ({
      id: w.id, name: w.name, phone: w.phone || "", rate: Number(w.base_rate) || 0,
      objectId: w.object_id || "", note: w.note || "", archived: !!w.archived, archivedDate: w.archived_date || "",
      permitExpiry: w.permit_expiry || "", contractExpiry: w.contract_expiry || "",
      rateCurrency: w.rate_currency === "CZK" ? "CZK" : "EUR",
    })),
    objects: (objects.data || []).map((o) => ({ id: o.id, name: o.name, billRate: billMap[o.id]?.rate || 0, billCur: billMap[o.id]?.cur || countryCur(o.country || "HR"), country: o.country || "HR" })),
    logs: (logs.data || []).map((l) => ({
      id: l.id, workerId: l.worker_id, objectId: l.object_id || "", date: l.work_date,
      from: l.from_t || "", to: l.to_t || "", hours: Number(l.hours), monthly: !!l.monthly, note: l.note || "",
      createdBy: l.created_by || "",
    })),
    payments: (payments.data || []).map((p) => ({
      id: p.id, workerId: p.worker_id || "", objectId: p.object_id || "", date: p.pay_date, type: p.type,
      amount: Number(p.amount), note: p.note || "", deduct: !!p.deduct,
      currency: p.currency || "EUR",
    })),
    assignments: (assignments.data || []).map((a) => ({ id: a.id, workerId: a.worker_id, objectId: a.object_id || "", from: a.from_date })),
    rateChanges: (rateChanges.data || []).map((r) => ({ id: r.id, workerId: r.worker_id, rate: Number(r.rate), from: r.from_date })),
    payouts: (payouts.data || []).map((p) => ({ id: p.id, workerId: p.worker_id, month: p.month, amount: Number(p.amount), amountKc: Number(p.amount_czk) || 0, paidAt: p.paid_at })),
    invoicePayments: invoicePayments.map((p) => ({ id: p.id, objectId: p.object_id, month: p.month, amount: Number(p.amount), date: p.pay_date, note: p.note || "" })),
    objectInvoices: objectInvoices.map((i) => ({
      id: i.id, objectId: i.object_id, number: i.number, period: i.period, issueDate: i.issue_date,
      hours: Number(i.hours) || 0, rate: Number(i.rate) || 0, amount: Number(i.amount) || 0,
      currency: i.currency === "CZK" ? "CZK" : "EUR", note: i.note || "",
    })),
    settings: settings || {},
    profiles: profiles.data || [],
    objectMembers: members,
    audit, trash,
  };
}

/* ================================================================== */
export default function App() {
  const { lang, setLang, t } = useLang();
  const [session, setSession] = useState(undefined);
  const [profile, setProfile] = useState(null);
  const [data, setData] = useState(null);
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [tab, setTab] = useState("radnici");
  const [openWorker, setOpenWorker] = useState(null);
  const [openObject, setOpenObject] = useState(null);
  const [adminPanel, setAdminPanel] = useState("");

  const admin = profile?.role === "admin";

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: sub } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  const reload = useCallback(async (prof) => {
    const p = prof || profile;
    if (!p) return;
    try { setData(await fetchAll(p.role === "admin")); setErr(""); }
    catch (e) { setErr("Greška pri učitavanju: " + (e.message || e)); }
  }, [profile]);

  useEffect(() => {
    if (!session) { setProfile(null); setData(null); return; }
    (async () => {
      const { data: p, error } = await supabase.from("profiles").select("*").eq("id", session.user.id).single();
      if (error) { setErr("Ne mogu učitati profil: " + error.message); return; }
      setProfile(p);
      await reload(p);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [session]);

  useEffect(() => {
    const onFocus = () => { if (profile) reload(); };
    window.addEventListener("focus", onFocus);
    return () => window.removeEventListener("focus", onFocus);
  }, [profile, reload]);

  const act = async (fn, auditText) => {
    setBusy(true);
    try {
      await fn();
      if (auditText) await supabase.from("audit_log").insert({ user_id: session.user.id, user_name: profile?.name || "", action: auditText });
      await reload();
      setErr("");
      return true;
    } catch (e) { setErr("Greška: " + (e.message || e)); return false; }
    finally { setBusy(false); }
  };
  const ins = (table, row) => supabase.from(table).insert(row).then(({ error }) => { if (error) throw error; });
  const upd = (table, id, patch) => supabase.from(table).update(patch).eq("id", id).then(({ error }) => { if (error) throw error; });
  const softDel = (table, id) => upd(table, id, { deleted_at: new Date().toISOString() });
  const guardPaid = (workerId, dateOrMonth) => {
    const mo = dateOrMonth.length === 7 ? dateOrMonth : monthKey(dateOrMonth);
    if (paidFor(data, workerId, mo)) throw new Error(`Mjesec ${mo} je označen kao ISPLAĆEN i zaključan za tog radnika. Admin ga može otključati u Obračunu.`);
  };

  const api = {
    admin, uid: () => session?.user?.id, settings: data?.settings || {}, t, lang, setLang,
    addWorker: (f) => act(async () => {
      const { data: w, error } = await supabase.from("workers")
        .insert({ name: f.name, phone: f.phone, base_rate: parseNum(f.rate) || 0, rate_currency: f.rateCur || "EUR", object_id: f.objectId || null, note: f.note,
          permit_expiry: f.permitExpiry || null, contract_expiry: f.contractExpiry || null, created_by: session.user.id })
        .select().single();
      if (error) throw error;
      if (f.objectId) await ins("assignments", { worker_id: w.id, object_id: f.objectId, from_date: todayISO(), created_by: session.user.id });
    }, `Dodao radnika: ${f.name}`),
    updWorker: (id, f, name) => act(() => upd("workers", id, {
      name: f.name, phone: f.phone, base_rate: parseNum(f.rate) || 0, rate_currency: f.rateCur || "EUR", object_id: f.objectId || null, note: f.note,
      permit_expiry: f.permitExpiry || null, contract_expiry: f.contractExpiry || null,
    }), `Uredio podatke radnika: ${name}`),
    delWorker: (w) => act(() => softDel("workers", w.id), `Obrisao radnika: ${w.name}`),
    archiveWorker: (w, on) => act(() => upd("workers", w.id, {
      archived: on, archived_date: on ? todayISO() : w.archivedDate || null, object_id: on ? null : (w.objectId || null),
    }), on ? `Označio da je ${w.name} završio s radom` : `Vratio ${w.name} u aktivne`),
    transfer: (w, objectId, from, targetName) => act(async () => {
      await upd("workers", w.id, { object_id: objectId || null });
      await ins("assignments", { worker_id: w.id, object_id: objectId || null, from_date: from, created_by: session.user.id });
    }, `Prebacio ${w.name} na: ${targetName} (od ${fmtDate(from)})`),
    delAssignment: (a, wName) => act(() => softDel("assignments", a.id), `Obrisao premještaj radnika ${wName}`),
    changeRate: (w, rate, from) => act(() => ins("rate_changes", { worker_id: w.id, rate, from_date: from, created_by: session.user.id }),
      `Promijenio satnicu: ${w.name} → ${eur(rate)}/h od ${fmtDate(from)}`),
    delRateChange: (r, wName) => act(() => softDel("rate_changes", r.id), `Obrisao promjenu satnice radnika ${wName}`),
    addObject: async (name, country = "HR") => {
      const { data: o, error } = await supabase.from("objects").insert({ name, country, created_by: session.user.id }).select().single();
      if (error) { setErr("Greška: " + error.message); return null; }
      await supabase.from("audit_log").insert({ user_id: session.user.id, user_name: profile?.name || "", action: `Dodao objekt: ${name} (${country})` });
      await reload();
      return o.id;
    },
    setObjectCountry: (o, country) => act(() => upd("objects", o.id, { country }),
      `Premjestio objekt ${o.name} u: ${COUNTRY_NAME[country] || country}`),
    delObject: (o) => act(() => softDel("objects", o.id), `Obrisao objekt: ${o.name}`),
    setBillRate: (o, rate, cur) => act(() => supabase.from("object_billing").upsert({ object_id: o.id, bill_rate: rate, bill_currency: cur || "EUR" })
      .then(({ error }) => { if (error) throw error; }), `Promijenio naplatu objekta ${o.name} na ${money(rate, cur)}/h`),
    addLog: (l, wName, objName) => act(() => { guardPaid(l.workerId, l.date); return ins("work_logs", {
      worker_id: l.workerId, object_id: l.objectId || null, work_date: l.date,
      from_t: l.from, to_t: l.to, hours: l.hours, monthly: !!l.monthly, note: l.note || "", created_by: session.user.id,
    }); }, `Upisao sate: ${wName} ${fmtH(l.hours)}${objName ? " (" + objName + ")" : ""} (${fmtDate(l.date)})`),
    delLog: (l, wName) => act(() => { guardPaid(l.workerId, l.date); return softDel("work_logs", l.id); },
      `Obrisao sate: ${wName} ${fmtH(l.hours)} (${fmtDate(l.date)})`),
    updLog: (l, patch, wName) => act(() => {
      guardPaid(l.workerId, l.date);
      guardPaid(l.workerId, patch.date);
      return upd("work_logs", l.id, { work_date: patch.date, from_t: patch.from, to_t: patch.to, hours: patch.hours });
    }, `Ispravio sate: ${wName} ${fmtH(patch.hours)} (${fmtDate(l.date)} → ${fmtDate(patch.date)})`),
    copyDay: (object, fromDate, toDate) => act(async () => {
      const src = data.logs.filter((l) => l.objectId === object.id && l.date === fromDate && !l.monthly);
      if (!src.length) throw new Error(`Za ${fmtDate(fromDate)} nema upisa na ovom objektu.`);
      for (const l of src) {
        guardPaid(l.workerId, toDate);
        await ins("work_logs", { worker_id: l.workerId, object_id: object.id, work_date: toDate,
          from_t: l.from, to_t: l.to, hours: l.hours, note: l.note || "", created_by: session.user.id });
      }
    }, `Kopirao dan ${fmtDate(fromDate)} → ${fmtDate(toDate)} (${object.name})`),
    addPayment: (p, name) => act(() => { if (p.workerId) guardPaid(p.workerId, p.date); return ins("payments", {
      worker_id: p.workerId || null, object_id: p.objectId || null, pay_date: p.date, type: p.type, amount: p.amount, note: p.note, deduct: p.deduct,
      currency: p.currency || "EUR", created_by: session.user.id,
    }); }, `Upisao ${TYPE_LABEL[p.type].toLowerCase()}: ${name} ${money(p.amount, p.currency)}`),
    delPayment: (p, name) => act(() => { if (p.workerId) guardPaid(p.workerId, p.date); return softDel("payments", p.id); },
      `Obrisao ${(TYPE_LABEL[p.type] || p.type).toLowerCase()}: ${name} ${eur(p.amount)}`),
    markPaid: (w, mo, amount, amountKc) => act(() => ins("payouts", { worker_id: w.id, month: mo, amount, amount_czk: amountKc || 0, created_by: session.user.id }),
      `Označio ISPLAĆENO: ${w.name} za ${mo} (${[amount ? eur(amount) : "", amountKc ? czk(amountKc) : ""].filter(Boolean).join(" + ") || eur(0)})`),
    unmarkPaid: (payout, wName) => act(() => softDel("payouts", payout.id), `Otključao isplatu: ${wName} za ${payout.month}`),
    addInvoicePayment: (o, mo, amount) => act(() => ins("invoice_payments", { object_id: o.id, month: mo, amount }),
      `Upisao uplatu od ${o.name}: ${eur(amount)} (${mo})`),
    delInvoicePayment: (p, oName) => act(() => softDel("invoice_payments", p.id), `Obrisao uplatu od ${oName}: ${eur(p.amount)}`),
    addObjectInvoice: (inv, oName) => act(() => ins("object_invoices", {
      object_id: inv.objectId, number: inv.number, period: inv.period, issue_date: inv.issueDate,
      hours: inv.hours, rate: inv.rate, amount: inv.amount, currency: inv.currency, note: inv.note || "",
      created_by: session.user.id,
    }), `Generirao fakturu ${inv.number} za ${oName}: ${money(inv.amount, inv.currency)}`),
    delObjectInvoice: (inv, oName) => act(() => softDel("object_invoices", inv.id), `Obrisao fakturu ${inv.number} (${oName})`),
    saveSettings: (s) => act(() => supabase.from("settings").upsert({ id: 1, ...s }).then(({ error }) => { if (error) throw error; }),
      "Uredio podatke firme"),
    restore: (t) => act(() => upd(t.table, t.row.id, { deleted_at: null }), `Vratio iz koša: ${t.label}`),
    setRole: (p, role) => act(() => upd("profiles", p.id, { role }), `Promijenio ulogu: ${p.name} → ${role === "admin" ? "admin" : "zaposlenik"}`),
    toggleMember: (object, p, on) => act(() =>
      (on ? supabase.from("object_members").insert({ object_id: object.id, user_id: p.id })
          : supabase.from("object_members").delete().eq("object_id", object.id).eq("user_id", p.id)
      ).then(({ error }) => { if (error) throw error; }),
      on ? `Dodijelio objekt ${object.name} zaposleniku ${p.name}` : `Maknuo objekt ${object.name} zaposleniku ${p.name}`),
    logDoc: (text) => supabase.from("audit_log").insert({ user_id: session.user.id, user_name: profile?.name || "", action: text }),
  };

  if (session === undefined) return <Center>Učitavam…</Center>;
  if (!session) return <Login lang={lang} setLang={setLang} t={t} />;
  if (!data) return <Center>{err || "Učitavam podatke…"}</Center>;

  const worker = data.workers.find((w) => w.id === openWorker);
  const object = data.objects.find((o) => o.id === openObject);
  const detailOpen = worker || object;

  return (
    <div style={{ minHeight: "100vh", background: S.bg, color: S.ink, fontFamily: "'Segoe UI', system-ui, sans-serif", paddingBottom: 84 }}>
      <style>{`
        * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
        input, select, button, textarea { font-family: inherit; font-size: 15px; }
        input, select, textarea { width: 100%; padding: 11px 12px; border: 1px solid ${S.line}; border-radius: 10px; background: #fff; color: ${S.ink}; outline: none; }
        input:focus, select:focus, textarea:focus { border-color: ${S.green}; }
        .num { font-variant-numeric: tabular-nums; }
        button:active { transform: scale(.985); }
      `}</style>

      <Header profile={profile} admin={admin} onLogout={() => supabase.auth.signOut()} lang={lang} setLang={setLang} t={t} />

      <div style={{ maxWidth: 560, margin: "0 auto", padding: "0 14px" }}>
        {err && (
          <div style={{ background: S.redSoft, color: S.red, borderRadius: 10, padding: "9px 12px", marginBottom: 12, fontSize: 13.5, fontWeight: 600 }}>
            {err} <button onClick={() => setErr("")} style={{ background: "none", border: "none", color: S.red, fontWeight: 800, cursor: "pointer", float: "right" }}>✕</button>
          </div>
        )}
        {busy && <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 8 }}>Spremam…</div>}

        {admin && !detailOpen && <AdminPanels data={data} api={api} panel={adminPanel} setPanel={setAdminPanel} />}

        {worker ? (
          <WorkerDetail worker={worker} data={data} api={api} onBack={() => setOpenWorker(null)} />
        ) : object ? (
          <ObjectDetail object={object} data={data} api={api} onBack={() => setOpenObject(null)} />
        ) : (
          <>
            {tab === "radnici" && <WorkersTab data={data} api={api} onOpen={setOpenWorker} onOpenObject={setOpenObject} />}
            {tab === "objekti" && <ObjectsTab data={data} api={api} onOpenObject={setOpenObject} />}
            {tab === "imenik" && <DirectoryTab data={data} api={api} onOpen={setOpenWorker} />}
            {tab === "sati" && <HoursTab data={data} api={api} />}
            {tab === "isplate" && <PaymentsTab data={data} api={api} />}
            {tab === "obracun" && <ReportTab data={data} api={api} admin={admin} />}
          </>
        )}
      </div>

      {!detailOpen && <TabBar tab={tab} setTab={setTab} t={t} />}
    </div>
  );
}

/* ================================================================== */
function Center({ children }) {
  return <div style={{ minHeight: "100vh", display: "grid", placeItems: "center", background: S.bg, color: S.sub, fontFamily: "system-ui", padding: 20, textAlign: "center" }}>{children}</div>;
}

function Login({ lang, setLang, t }) {
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const login = async () => {
    setBusy(true); setErr("");
    const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password: pass });
    setBusy(false);
    if (error) setErr(t("loginErr"));
  };
  return (
    <div style={{ minHeight: "100vh", background: S.bg, fontFamily: "'Segoe UI', system-ui, sans-serif", color: S.ink }}>
      <style>{`* { box-sizing: border-box; } input { width: 100%; padding: 12px; border: 1px solid ${S.line}; border-radius: 10px; outline: none; font-size: 15px; } input:focus { border-color: ${S.green}; }`}</style>
      <div style={{ background: S.ink, color: "#fff", padding: "26px 16px" }}>
        <div style={{ maxWidth: 420, margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 11, letterSpacing: 2, textTransform: "uppercase", color: "#9DB3A8" }}>{t("appTitle")}</div>
            <div style={{ fontSize: 22, fontWeight: 800, marginTop: 2 }}>{t("loginTitle")}</div>
          </div>
          <LangSwitch lang={lang} setLang={setLang} dark />
        </div>
      </div>
      <div style={{ maxWidth: 420, margin: "26px auto 0", padding: "0 16px" }}>
        <div style={{ background: "#fff", border: `1px solid ${S.line}`, borderRadius: 14, padding: 18 }}>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: S.sub, marginBottom: 5 }}>{t("email")}</div>
          <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="ime@firma.hr" autoComplete="username" />
          <div style={{ fontSize: 12.5, fontWeight: 600, color: S.sub, margin: "12px 0 5px" }}>{t("password")}</div>
          <input type="password" value={pass} onChange={(e) => setPass(e.target.value)} autoComplete="current-password" onKeyDown={(e) => e.key === "Enter" && login()} />
          {err && <div style={{ color: S.red, fontSize: 13, marginTop: 10 }}>{err}</div>}
          <button onClick={login} disabled={busy} style={{ width: "100%", background: S.green, color: "#fff", border: "none", borderRadius: 10, padding: 13, fontWeight: 700, cursor: "pointer", marginTop: 14, fontSize: 15, opacity: busy ? 0.7 : 1 }}>
            {busy ? t("loginBusy") : t("loginBtn")}
          </button>
          <div style={{ fontSize: 12.5, color: S.sub, marginTop: 14, lineHeight: 1.5 }}>
            {t("noAccount")}
          </div>
        </div>
      </div>
    </div>
  );
}

function Header({ profile, admin, onLogout, lang, setLang, t }) {
  return (
    <div style={{ background: S.ink, color: "#fff", padding: "16px 16px 12px", marginBottom: 16 }}>
      <div style={{ maxWidth: 560, margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <div>
          <div style={{ fontSize: 11, letterSpacing: 2, textTransform: "uppercase", color: "#9DB3A8" }}>{t("appTitle")}</div>
          <div style={{ fontSize: 20, fontWeight: 700, marginTop: 2 }}>{t("appSubtitle")}</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <LangSwitch lang={lang} setLang={setLang} dark />
          <button onClick={onLogout} style={{ background: "rgba(255,255,255,.1)", border: "1px solid rgba(255,255,255,.2)", borderRadius: 10, color: "#fff", padding: "7px 11px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>
            {admin ? "👑" : "👤"} {profile?.name || "?"} · {t("logout")}
          </button>
        </div>
      </div>
    </div>
  );
}
function TabBar({ tab, setTab, t }) {
  const tabs = [["radnici",t("tabRadnici"),"👷"],["objekti",t("tabObjekti"),"🏨"],["imenik",t("tabImenik"),"📇"],["sati",t("tabSati"),"🕒"],["isplate",t("tabIsplate"),"💶"],["obracun",t("tabObracun"),"🧾"]];
  return (
    <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: "#fff", borderTop: `1px solid ${S.line}`,
      display: "flex", justifyContent: "space-around", padding: "6px 4px calc(6px + env(safe-area-inset-bottom))", zIndex: 20 }}>
      {tabs.map(([id, label, icon]) => (
        <button key={id} onClick={() => setTab(id)} style={{ background: "none", border: "none", padding: "6px 8px", borderRadius: 10,
          color: tab === id ? S.green : S.sub, fontWeight: tab === id ? 700 : 500, fontSize: 11.5,
          display: "flex", flexDirection: "column", alignItems: "center", gap: 2, cursor: "pointer" }}>
          <span style={{ fontSize: 19 }}>{icon}</span>{label}
        </button>
      ))}
    </div>
  );
}
function Card({ children, style, onClick }) {
  return <div onClick={onClick} style={{ background: S.card, border: `1px solid ${S.line}`, borderRadius: 14, padding: 14, marginBottom: 12, ...style }}>{children}</div>;
}
function Btn({ children, onClick, kind = "primary", small, style, disabled }) {
  const base = {
    primary: { background: S.green, color: "#fff", border: "none" },
    ghost: { background: "#fff", color: S.ink, border: `1px solid ${S.line}` },
    danger: { background: "#fff", color: S.red, border: `1px solid ${S.redSoft}` },
    excel: { background: "#1D6F42", color: "#fff", border: "none" },
  }[kind];
  return (
    <button onClick={onClick} disabled={disabled} style={{ ...base, borderRadius: 10, padding: small ? "7px 12px" : "12px 16px",
      fontWeight: 600, cursor: "pointer", fontSize: small ? 13.5 : 15, whiteSpace: "nowrap", opacity: disabled ? 0.6 : 1, ...style }}>
      {children}
    </button>
  );
}
function Field({ label, children }) {
  return (
    <label style={{ display: "block", marginBottom: 10 }}>
      <div style={{ fontSize: 12.5, fontWeight: 600, color: S.sub, marginBottom: 5 }}>{label}</div>
      {children}
    </label>
  );
}
function Empty({ text }) { return <div style={{ textAlign: "center", color: S.sub, padding: "34px 10px", fontSize: 14.5 }}>{text}</div>; }
function Tag({ children, color, bg }) {
  return <span style={{ background: bg, color, borderRadius: 999, padding: "3px 10px", fontSize: 12, fontWeight: 700, whiteSpace: "nowrap" }}>{children}</span>;
}
function MiniBars({ items, fmt, color }) {
  const max = Math.max(...items.map((i) => i.value), 0.001);
  return (
    <div style={{ display: "flex", gap: 6, alignItems: "flex-end", padding: "6px 2px 0" }}>
      {items.map((i) => (
        <div key={i.label} style={{ flex: 1, textAlign: "center" }}>
          <div className="num" style={{ fontSize: 10, color: S.sub, marginBottom: 2 }}>{fmt(i.value)}</div>
          <div style={{ height: Math.max(3, Math.round(70 * i.value / max)), background: i.value >= 0 ? color : S.red, borderRadius: "6px 6px 0 0" }} />
          <div style={{ fontSize: 10.5, color: S.sub, marginTop: 3 }}>{i.label}</div>
        </div>
      ))}
    </div>
  );
}
/* Uređivanje postojećeg upisa sati (ispravak datuma/vremena) */
function LogEditor({ log, onSave, onCancel }) {
  const [f, setF] = useState({ date: log.date, from: log.from || "", to: log.to || "", hours: String(log.hours) });
  const save = () => {
    let hours = parseNum(f.hours);
    if (f.from && f.to && String(log.hours) === f.hours) hours = hoursBetween(f.from, f.to);
    if (!hours || hours <= 0 || !f.date) return;
    onSave({ date: f.date, from: f.from && f.to ? f.from : "", to: f.from && f.to ? f.to : "", hours: round2(hours) });
  };
  return (
    <div style={{ background: S.amberSoft, borderRadius: 10, padding: "8px 10px", margin: "4px 0 8px" }}>
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
        <input type="date" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} style={{ flex: "1 1 130px", padding: "8px" }} />
        {!log.monthly && (
          <>
            <input type="time" value={f.from} onChange={(e) => setF({ ...f, from: e.target.value })} style={{ flex: "1 1 80px", padding: "8px 6px" }} />
            <input type="time" value={f.to} onChange={(e) => setF({ ...f, to: e.target.value })} style={{ flex: "1 1 80px", padding: "8px 6px" }} />
          </>
        )}
        <input inputMode="decimal" value={f.hours} onChange={(e) => setF({ ...f, hours: e.target.value })} placeholder="sati" style={{ flex: "1 1 60px", padding: "8px" }} />
      </div>
      <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
        <Btn small onClick={save}>Spremi ispravak</Btn>
        <Btn small kind="ghost" onClick={onCancel}>Odustani</Btn>
      </div>
    </div>
  );
}

function ObjectSelect({ data, api, value, onChange }) {
  const [adding, setAdding] = useState(false);
  const [name, setName] = useState("");
  const commit = async () => {
    const n = name.trim();
    if (!n) { setAdding(false); setName(""); return; }
    const exists = data.objects.find((o) => o.name.toLowerCase() === n.toLowerCase());
    if (exists) onChange(exists.id);
    else { const id = await api.addObject(n); if (id) onChange(id); }
    setName(""); setAdding(false);
  };
  if (adding) return (
    <div style={{ display: "flex", gap: 8 }}>
      <input autoFocus value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && commit()} placeholder="Naziv objekta" />
      <Btn small onClick={commit}>Dodaj</Btn>
      <Btn small kind="ghost" onClick={() => { setAdding(false); setName(""); }}>✕</Btn>
    </div>
  );
  return (
    <select value={value} onChange={(e) => (e.target.value === "__new__" ? setAdding(true) : onChange(e.target.value))}>
      <option value="">— bez objekta —</option>
      {["HR", "CZ"].map((c) => {
        const objs = sortedObjects(data.objects.filter((o) => (o.country || "HR") === c));
        return objs.length ? (
          <optgroup key={c} label={`${FLAG[c]} ${COUNTRY_NAME[c]}`}>
            {objs.map((o) => <option key={o.id} value={o.id}>{o.name}</option>)}
          </optgroup>
        ) : null;
      })}
      <option value="__new__">＋ Dodaj novi objekt…</option>
    </select>
  );
}

/* ================================================================== */
/*  ADMIN PANELI                                                       */
/* ================================================================== */
function AdminPanels({ data, api, panel, setPanel }) {
  const [firm, setFirm] = useState({ company_name: "", address: "", oib: "", iban: "" });
  useEffect(() => setFirm({
    company_name: data.settings.company_name || "", address: data.settings.address || "",
    oib: data.settings.oib || "", iban: data.settings.iban || "",
  }), [data.settings]);

  return (
    <div style={{ marginBottom: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8,
        background: S.amberSoft, border: `1px solid #EBD9B4`, borderRadius: 12, padding: "9px 12px" }}>
        <span style={{ fontWeight: 700, color: S.amber, fontSize: 13.5 }}>👑 Admin</span>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {[["users","👥"],["audit","📜"],["trash",`🗑${data.trash.length ? data.trash.length : ""}`],["firm","⚙️"]].map(([id, label]) => (
            <button key={id} onClick={() => setPanel(panel === id ? "" : id)} style={{
              background: panel === id ? S.amber : "#fff", color: panel === id ? "#fff" : S.amber,
              border: `1px solid ${S.amber}`, borderRadius: 8, padding: "5px 10px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
              {label}
            </button>
          ))}
        </div>
      </div>

      {panel === "users" && (
        <Card style={{ marginTop: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>👥 Zaposlenici s pristupom</div>
          {data.profiles.map((p) => (
            <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderBottom: `1px solid ${S.line}` }}>
              <span style={{ flex: 1, fontWeight: 600, fontSize: 14 }}>{p.role === "admin" ? "👑 " : ""}{p.name || "(bez imena)"}</span>
              <select value={p.role} onChange={(e) => api.setRole(p, e.target.value)} style={{ width: "auto", padding: "6px 8px", fontSize: 13 }}>
                <option value="employee">Zaposlenik</option>
                <option value="admin">Admin</option>
              </select>
            </div>
          ))}
          <div style={{ fontSize: 12.5, color: S.sub, marginTop: 10 }}>
            Novog zaposlenika dodaješ u Supabaseu: Authentication → Users → Add user.
          </div>
        </Card>
      )}

      {panel === "audit" && (
        <Card style={{ marginTop: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>📜 Tko je što radio</div>
          {data.audit.length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>Još nema zabilježenih radnji.</div>
            : data.audit.map((a) => (
              <div key={a.id} style={{ padding: "6px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13 }}>
                <span className="num" style={{ color: S.sub }}>{fmtDT(a.at)}</span>{" · "}<b>{a.user_name}</b> — {a.action}
              </div>
            ))}
        </Card>
      )}

      {panel === "trash" && (
        <Card style={{ marginTop: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>🗑 Obrisano — može se vratiti</div>
          {data.trash.length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>Koš je prazan.</div>
            : data.trash.map((t) => (
              <div key={t.table + t.row.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderBottom: `1px solid ${S.line}` }}>
                <div style={{ flex: 1, fontSize: 13 }}>
                  <div style={{ fontWeight: 600 }}>{t.label}</div>
                  <div className="num" style={{ fontSize: 11.5, color: S.sub }}>{fmtDT(t.row.deleted_at)}</div>
                </div>
                <Btn small kind="ghost" onClick={() => api.restore(t)} style={{ color: S.green, fontWeight: 700 }}>↩ Vrati</Btn>
              </div>
            ))}
        </Card>
      )}

      {panel === "firm" && (
        <Card style={{ marginTop: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>⚙️ Podaci firme (za specifikacije i obračune)</div>
          <Field label="Naziv firme"><input value={firm.company_name} onChange={(e) => setFirm({ ...firm, company_name: e.target.value })} placeholder="npr. Moja Firma d.o.o." /></Field>
          <Field label="Adresa"><input value={firm.address} onChange={(e) => setFirm({ ...firm, address: e.target.value })} /></Field>
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}><Field label="OIB"><input value={firm.oib} onChange={(e) => setFirm({ ...firm, oib: e.target.value })} /></Field></div>
            <div style={{ flex: 1.4 }}><Field label="IBAN"><input value={firm.iban} onChange={(e) => setFirm({ ...firm, iban: e.target.value })} /></Field></div>
          </div>
          <Btn small onClick={() => api.saveSettings(firm)}>Spremi</Btn>
        </Card>
      )}
    </div>
  );
}

/* ================================================================== */
/*  RADNICI + OBJEKTI + upozorenja na dokumente                        */
/* ================================================================== */
function expiryWarnings(workers) {
  const soon = addDaysISO(todayISO(), 30);
  const out = [];
  workers.filter((w) => !w.archived).forEach((w) => {
    [["permitExpiry", "radna dozvola"], ["contractExpiry", "ugovor"]].forEach(([k, label]) => {
      const d = w[k];
      if (d && d <= soon) out.push({ w, label, date: d, past: d < todayISO() });
    });
  });
  return out.sort((a, b) => a.date.localeCompare(b.date));
}

function WorkersTab({ data, api, onOpen, onOpenObject }) {
  const [adding, setAdding] = useState(false);
  const [showObjects, setShowObjects] = useState(true);
  const [form, setForm] = useState({ name: "", phone: "", rate: "", rateCur: "EUR", objectId: "", note: "", permitExpiry: "", contractExpiry: "" });
  const [newObj, setNewObj] = useState("");
  const [confirmObj, setConfirmObj] = useState(null);

  const objName = (id) => data.objects.find((o) => o.id === id)?.name || "";
  const mk = curMonth();
  const warns = expiryWarnings(data.workers);

  const addWorker = async () => {
    if (!form.name.trim()) return;
    if (await api.addWorker(form)) { setForm({ name: "", phone: "", rate: "", rateCur: "EUR", objectId: "", note: "", permitExpiry: "", contractExpiry: "" }); setAdding(false); }
  };
  const addObject = async () => {
    const n = newObj.trim(); if (!n) return;
    if (data.objects.find((o) => o.name.toLowerCase() === n.toLowerCase())) { setNewObj(""); return; }
    await api.addObject(n); setNewObj("");
  };
  const delObject = (o) => {
    if (confirmObj !== o.id) { setConfirmObj(o.id); return; }
    setConfirmObj(null); api.delObject(o);
  };

  const actives = sortedWorkers(data.workers.filter((w) => !w.archived));
  const formerWorkers = sortedWorkers(data.workers.filter((w) => w.archived));
  const objectsSorted = sortedObjects(data.objects);

  return (
    <>
      {warns.length > 0 && (
        <Card style={{ background: S.redSoft, borderColor: "#EED0C8" }}>
          <div style={{ fontWeight: 700, color: S.red, marginBottom: 6 }}>⚠️ Istek dokumenata</div>
          {warns.map((x, i) => (
            <div key={i} onClick={() => onOpen(x.w.id)} style={{ fontSize: 13.5, padding: "4px 0", cursor: "pointer" }}>
              <b>{x.w.name}</b> — {x.label} {x.past ? "ISTEKAO" : "istječe"} <span className="num">{fmtDate(x.date)}</span>
            </div>
          ))}
        </Card>
      )}

      <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA" }}>
        <div onClick={() => setShowObjects(!showObjects)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
          <div style={{ fontWeight: 700, color: S.blue }}>🏨 Objekti ({data.objects.length})</div>
          <span style={{ color: S.blue, fontWeight: 700 }}>{showObjects ? "▲" : "▼"}</span>
        </div>
        {showObjects && (
          <div style={{ marginTop: 10 }}>
            {objectsSorted.map((o) => {
              const h = data.logs.filter((l) => l.objectId === o.id && monthKey(l.date) === mk).reduce((s, l) => s + l.hours, 0);
              return (
                <div key={o.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: `1px solid #CBDCEA` }}>
                  <div onClick={() => onOpenObject(o.id)} style={{ flex: 1, cursor: "pointer" }}>
                    <span style={{ fontWeight: 700, fontSize: 15 }}>{FLAG[o.country || "HR"]} {o.name} <span style={{ color: S.blue }}>›</span></span>
                    <div className="num" style={{ fontSize: 12.5, color: S.sub }}>
                      {fmtH(round2(h))} ovaj mjesec
                      {api.admin && o.billRate ? <span style={{ color: S.amber, fontWeight: 700 }}> · naplata {money(o.billRate, o.billCur)}/h</span> : ""}
                    </div>
                  </div>
                  {api.admin && (
                    <button onClick={() => delObject(o)} style={{
                      background: confirmObj === o.id ? S.red : "none", color: confirmObj === o.id ? "#fff" : S.red,
                      border: "none", borderRadius: 8, cursor: "pointer",
                      fontSize: confirmObj === o.id ? 12.5 : 15, fontWeight: 700, padding: confirmObj === o.id ? "5px 10px" : 4 }}>
                      {confirmObj === o.id ? "Potvrdi brisanje" : "✕"}
                    </button>
                  )}
                </div>
              );
            })}
            <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
              <input value={newObj} onChange={(e) => setNewObj(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addObject()} placeholder="Naziv novog objekta…" />
              <Btn small onClick={addObject}>Dodaj</Btn>
            </div>
          </div>
        )}
      </Card>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <div style={{ fontSize: 17, fontWeight: 700 }}>{api.t("workersTitle")} ({actives.length})</div>
        <Btn small onClick={() => setAdding(!adding)}>{adding ? api.t("close") : api.t("addWorker")}</Btn>
      </div>

      {adding && (
        <Card>
          <Field label={api.t("fullName")}><input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="npr. Ivan Horvat" /></Field>
          <Field label={api.t("phone")}><input value={form.phone} inputMode="tel" onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
          <Field label={`Satnica (${form.rateCur === "CZK" ? "Kč" : "€"} / sat) — nije obavezno`}>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <input value={form.rate} inputMode="decimal" onChange={(e) => setForm({ ...form, rate: e.target.value })} placeholder={form.rateCur === "CZK" ? "npr. 150" : "npr. 6.50"} />
              <CurChips small value={form.rateCur} onChange={(v) => setForm({ ...form, rateCur: v })} />
            </div>
          </Field>
          <Field label="Glavni objekt"><ObjectSelect data={data} api={api} value={form.objectId} onChange={(v) => {
            const ob = data.objects.find((o) => o.id === v);
            setForm({ ...form, objectId: v, rateCur: ob ? countryCur(ob.country) : form.rateCur });
          }} /></Field>
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}><Field label="Istek radne dozvole"><input type="date" value={form.permitExpiry} onChange={(e) => setForm({ ...form, permitExpiry: e.target.value })} /></Field></div>
            <div style={{ flex: 1 }}><Field label="Istek ugovora"><input type="date" value={form.contractExpiry} onChange={(e) => setForm({ ...form, contractExpiry: e.target.value })} /></Field></div>
          </div>
          <Field label={api.t("note") + " (OIB…)"}><textarea rows={2} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
          <Btn onClick={addWorker} style={{ width: "100%" }}>{api.t("saveWorker")}</Btn>
        </Card>
      )}

      {actives.length === 0 && !adding && <Empty text={api.t("noWorkersYet")} />}

      {actives.map((w) => {
        const h = data.logs.filter((l) => l.workerId === w.id && monthKey(l.date) === mk).reduce((s, l) => s + l.hours, 0);
        return (
          <Card key={w.id} style={{ cursor: "pointer" }}>
            <div onClick={() => onOpen(w.id)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 15.5 }}>{w.name}</div>
                <div style={{ fontSize: 13, color: S.sub, marginTop: 2 }}>{objName(w.objectId) || "—"} · {w.phone || "bez broja"}</div>
              </div>
              <div style={{ textAlign: "right" }}>
                {rateNow(data, w) > 0
                  ? <div className="num" style={{ fontWeight: 700, color: S.green }}>{money(rateNow(data, w), wCur(w))}/h</div>
                  : <div style={{ fontWeight: 700, color: S.amber, fontSize: 12.5 }}>{api.t("withoutRate")}</div>}
                <div className="num" style={{ fontSize: 12.5, color: S.sub }}>{fmtH(round2(h))} {api.t("thisMonth")}</div>
              </div>
            </div>
          </Card>
        );
      })}

      {data.workers.some((w) => w.archived) && (
        <>
          <div style={{ fontSize: 15, fontWeight: 700, color: S.sub, margin: "18px 0 8px" }}>📁 {api.t("formerWorkers")} ({formerWorkers.length})</div>
          {formerWorkers.map((w) => (
            <Card key={w.id} style={{ cursor: "pointer", opacity: 0.8 }}>
              <div onClick={() => onOpen(w.id)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 15 }}>{w.name}</div>
                  <div style={{ fontSize: 13, color: S.sub, marginTop: 2 }}>{w.phone || "bez broja"}{w.archivedDate ? " · završio " + fmtDate(w.archivedDate) : ""}</div>
                </div>
                <Tag color={S.sub} bg="#EEF0ED">arhiva</Tag>
              </div>
            </Card>
          ))}
        </>
      )}
    </>
  );
}

/* ================================================================== */
/*  OBJEKTI — pregled po objektu                                       */
/* ================================================================== */
function ObjectsTab({ data, api, onOpenObject }) {
  const mk = curMonth();
  const [newObj, setNewObj] = useState("");
  const [country, setCountry] = useState("HR"); // aktivni filter/tab — i zadana država za novi objekt
  const addObject = async () => {
    const n = newObj.trim(); if (!n) return;
    if (data.objects.find((o) => o.name.toLowerCase() === n.toLowerCase())) { setNewObj(""); return; }
    await api.addObject(n, country); setNewObj("");
  };
  const list = sortedObjects(data.objects);
  const countHR = list.filter((o) => (o.country || "HR") === "HR").length;
  const countCZ = list.filter((o) => o.country === "CZ").length;
  const shown = list.filter((o) => (o.country || "HR") === country);
  const renderObject = (o) => {
        const logs = data.logs.filter((l) => l.objectId === o.id && monthKey(l.date) === mk);
        const hrs = round2(logs.reduce((s, l) => s + l.hours, 0));
        const workerIds = new Set([
          ...data.workers.filter((w) => !w.archived && w.objectId === o.id).map((w) => w.id),
          ...logs.map((l) => l.workerId),
        ]);
        const revenue = round2(hrs * (o.billRate || 0));
        return (
          <Card key={o.id} style={{ cursor: "pointer" }} onClick={() => onOpenObject(o.id)}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 15.5 }}>{o.name} <span style={{ color: S.blue }}>›</span></div>
                <div className="num" style={{ fontSize: 13, color: S.sub, marginTop: 2 }}>
                  {workerIds.size} radnika · {fmtH(hrs)} ovaj mjesec
                </div>
              </div>
              {api.admin && o.billRate > 0 && (
                <div className="num" style={{ textAlign: "right" }}>
                  <div style={{ fontWeight: 700, color: S.amber }}>{money(revenue, o.billCur)}</div>
                  <div style={{ fontSize: 11.5, color: S.sub }}>naplata ovaj mj.</div>
                </div>
              )}
            </div>
          </Card>
        );
  };

  return (
    <>
      <div style={{ fontSize: 17, fontWeight: 700, marginBottom: 10 }}>🏨 {api.t("objectsTitle")} ({list.length})</div>
      <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
        {["HR", "CZ"].map((c) => (
          <button key={c} onClick={() => setCountry(c)} style={{ flex: 1, padding: "10px 8px", borderRadius: 10, fontWeight: 700, fontSize: 13.5, cursor: "pointer",
            background: country === c ? S.blue : "#fff", color: country === c ? "#fff" : S.sub, border: `1px solid ${country === c ? S.blue : S.line}` }}>
            {FLAG[c]} {COUNTRY_NAME[c]} ({c === "HR" ? countHR : countCZ})
          </button>
        ))}
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
        <input value={newObj} onChange={(e) => setNewObj(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addObject()} placeholder={`${api.t("newObjectPh")} ${COUNTRY_NAME[country]}…`} />
        <Btn small onClick={addObject}>{api.t("addBtn")}</Btn>
      </div>
      {shown.length === 0 ? <Empty text={`Još nema objekata u: ${COUNTRY_NAME[country]}.`} /> : shown.map(renderObject)}
    </>
  );
}

/* ================================================================== */
/*  IMENIK                                                             */
/* ================================================================== */
function DirectoryTab({ data, api, onOpen }) {
  const [q, setQ] = useState("");
  const list = [...data.workers].sort((a, b) => a.name.localeCompare(b.name, "hr"))
    .filter((w) => { const t = q.trim().toLowerCase(); return !t || w.name.toLowerCase().includes(t) || (w.phone || "").includes(t); });
  const active = data.workers.filter((w) => !w.archived).length;
  const former = data.workers.length - active;
  const groups = [];
  let last = "";
  list.forEach((w) => {
    const letter = (w.name.trim()[0] || "#").toUpperCase();
    if (letter !== last) { groups.push({ letter, workers: [w] }); last = letter; }
    else groups[groups.length - 1].workers.push(w);
  });
  return (
    <>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <div style={{ fontSize: 17, fontWeight: 700 }}>📇 {api.t("directoryTitle")}</div>
        <div style={{ fontSize: 12.5, color: S.sub }}>
          <span style={{ color: S.green, fontWeight: 700 }}>{active} aktivnih</span>{former > 0 && <> · {former} bivših</>}
        </div>
      </div>
      <input value={q} onChange={(e) => setQ(e.target.value)} placeholder={api.t("searchPh")} style={{ marginBottom: 12 }} />
      {list.length === 0 ? <Empty text="Nema rezultata." /> : groups.map((g) => (
        <div key={g.letter}>
          <div style={{ fontSize: 13, fontWeight: 800, color: S.blue, padding: "6px 4px 4px" }}>{g.letter}</div>
          <Card style={{ padding: "4px 14px" }}>
            {g.workers.map((w, i) => (
              <div key={w.id} onClick={() => onOpen(w.id)} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 0", cursor: "pointer",
                borderBottom: i < g.workers.length - 1 ? `1px solid ${S.line}` : "none", opacity: w.archived ? 0.75 : 1 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 15 }}>{w.name}</div>
                  <div style={{ fontSize: 12.5, color: S.sub, marginTop: 1 }}>
                    {w.phone ? <a href={"tel:" + w.phone} onClick={(e) => e.stopPropagation()} style={{ color: S.sub }}>📞 {w.phone}</a> : "bez broja"}
                  </div>
                </div>
                {w.archived
                  ? <Tag color={S.sub} bg="#EEF0ED">Završio{w.archivedDate ? " " + fmtDate(w.archivedDate) : ""}</Tag>
                  : <Tag color={S.green} bg={S.greenSoft}>{api.t("active")}</Tag>}
              </div>
            ))}
          </Card>
        </div>
      ))}
    </>
  );
}

/* ================================================================== */
/*  DETALJ RADNIKA                                                     */
/* ================================================================== */
function WorkerDetail({ worker, data, api, onBack }) {
  const [edit, setEdit] = useState(false);
  const [form, setForm] = useState({ ...worker, rate: String(worker.rate || ""), rateCur: wCur(worker) });
  const [confirmRemove, setConfirmRemove] = useState(false);
  const [moveObj, setMoveObj] = useState(worker.objectId || "");
  const [moveDate, setMoveDate] = useState(todayISO());
  const [rateNew, setRateNew] = useState("");
  const [rateFrom, setRateFrom] = useState(todayISO());
  const [docs, setDocs] = useState(null);
  const [docErr, setDocErr] = useState("");
  const [viewer, setViewer] = useState(null);
  const fileRef = useRef(null);

  const objName = (id) => data.objects.find((o) => o.id === id)?.name || "";

  const loadDocs = useCallback(async () => {
    const { data: files, error } = await supabase.storage.from("docs").list(worker.id, { limit: 100 });
    if (error) { setDocErr(error.message); return; }
    setDocs(files || []);
  }, [worker.id]);
  useEffect(() => { loadDocs(); setMoveObj(worker.objectId || ""); }, [worker.id, worker.objectId, loadDocs]);

  const uploadDoc = async (e) => {
    const files = Array.from(e.target.files || []);
    setDocErr("");
    for (const f of files) {
      if (f.size > 10 * 1024 * 1024) { setDocErr(`"${f.name}" je veća od 10 MB.`); continue; }
      const { error } = await supabase.storage.from("docs").upload(`${worker.id}/${Date.now()}_${f.name}`, f);
      if (error) setDocErr(error.message);
    }
    await api.logDoc(`Dodao dokument radniku ${worker.name}`);
    await loadDocs();
    e.target.value = "";
  };
  const openDoc = async (f) => {
    const { data: s } = await supabase.storage.from("docs").createSignedUrl(`${worker.id}/${f.name}`, 300);
    if (s?.signedUrl) setViewer({ name: f.name, url: s.signedUrl, image: /\.(jpe?g|png|gif|webp)$/i.test(f.name) });
  };
  const delDoc = async (f) => {
    const { error } = await supabase.storage.from("docs").remove([`${worker.id}/${f.name}`]);
    if (error) { setDocErr("Brisanje dokumenata može samo admin."); return; }
    await api.logDoc(`Obrisao dokument radnika ${worker.name}`);
    setViewer(null);
    await loadDocs();
  };

  const save = async () => { if (await api.updWorker(worker.id, form, form.name)) setEdit(false); };
  const remove = async () => {
    if (!confirmRemove) { setConfirmRemove(true); return; }
    if (await api.delWorker(worker)) onBack();
  };
  const transfer = () => {
    if ((moveObj || "") === (worker.objectId || "")) return;
    api.transfer(worker, moveObj, moveDate, objName(moveObj) || "bez objekta");
  };
  const changeRate = () => {
    const r = parseNum(rateNew);
    if (!r || r <= 0) return;
    api.changeRate(worker, round2(r), rateFrom);
    setRateNew("");
  };

  const logs = data.logs.filter((l) => l.workerId === worker.id).sort((a, b) => b.date.localeCompare(a.date)).slice(0, 8);
  const myAssignments = data.assignments.filter((a) => a.workerId === worker.id).sort((a, b) => b.from.localeCompare(a.from));
  const myRates = data.rateChanges.filter((r) => r.workerId === worker.id).sort((a, b) => a.from.localeCompare(b.from));

  return (
    <>
      <button onClick={onBack} style={{ background: "none", border: "none", color: S.green, fontWeight: 700, fontSize: 14.5, padding: "2px 0 12px", cursor: "pointer" }}>← Natrag</button>

      <Card>
        {!edit ? (
          <>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ fontSize: 19, fontWeight: 800 }}>{worker.name}</div>
                <div style={{ color: S.sub, fontSize: 13.5, marginTop: 3 }}>{worker.archived ? "📁 Bivši radnik · " : ""}{objName(worker.objectId) || "Bez objekta"}</div>
              </div>
              {rateNow(data, worker) > 0
                ? <Tag color={S.green} bg={S.greenSoft}>{money(rateNow(data, worker), wCur(worker))}/h</Tag>
                : <Tag color={S.amber} bg={S.amberSoft}>bez satnice</Tag>}
            </div>
            <div style={{ marginTop: 10, fontSize: 14.5 }}>📞 {worker.phone ? <a href={"tel:" + worker.phone} style={{ color: S.ink }}>{worker.phone}</a> : "—"}</div>
            {(worker.permitExpiry || worker.contractExpiry) && (
              <div className="num" style={{ marginTop: 6, fontSize: 13, color: S.sub }}>
                {worker.permitExpiry && <>Radna dozvola do {fmtDate(worker.permitExpiry)}</>}
                {worker.permitExpiry && worker.contractExpiry && " · "}
                {worker.contractExpiry && <>Ugovor do {fmtDate(worker.contractExpiry)}</>}
              </div>
            )}
            {worker.note && <div style={{ marginTop: 8, fontSize: 13.5, color: S.sub, whiteSpace: "pre-wrap" }}>{worker.note}</div>}
            <div style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap" }}>
              <Btn small kind="ghost" onClick={() => { setForm({ ...worker, rate: String(worker.rate || ""), rateCur: wCur(worker) }); setEdit(true); }}>Uredi</Btn>
              <Btn small kind="ghost" onClick={() => api.archiveWorker(worker, !worker.archived)} style={{ color: worker.archived ? S.green : S.amber, fontWeight: 700 }}>
                {worker.archived ? "↩ Vrati u aktivne" : "📁 Završio s radom"}
              </Btn>
              {api.admin ? (
                <>
                  <Btn small kind={confirmRemove ? "primary" : "danger"} onClick={remove} style={confirmRemove ? { background: S.red } : undefined}>
                    {confirmRemove ? "Potvrdi brisanje" : "Obriši"}
                  </Btn>
                  {confirmRemove && <Btn small kind="ghost" onClick={() => setConfirmRemove(false)}>Odustani</Btn>}
                </>
              ) : <span style={{ fontSize: 12, color: S.sub, alignSelf: "center" }}>🔒 brisanje može samo admin</span>}
            </div>
          </>
        ) : (
          <>
            <Field label="Ime i prezime"><input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
            <Field label="Broj telefona"><input value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
            <Field label={`Satnica (${form.rateCur === "CZK" ? "Kč" : "€"} / sat)`}>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input value={form.rate} inputMode="decimal" onChange={(e) => setForm({ ...form, rate: e.target.value })} />
                <CurChips small value={form.rateCur} onChange={(v) => setForm({ ...form, rateCur: v })} />
              </div>
            </Field>
            <Field label="Glavni objekt"><ObjectSelect data={data} api={api} value={form.objectId || ""} onChange={(v) => setForm({ ...form, objectId: v })} /></Field>
            <div style={{ display: "flex", gap: 10 }}>
              <div style={{ flex: 1 }}><Field label="Istek radne dozvole"><input type="date" value={form.permitExpiry || ""} onChange={(e) => setForm({ ...form, permitExpiry: e.target.value })} /></Field></div>
              <div style={{ flex: 1 }}><Field label="Istek ugovora"><input type="date" value={form.contractExpiry || ""} onChange={(e) => setForm({ ...form, contractExpiry: e.target.value })} /></Field></div>
            </div>
            <Field label="Napomena"><textarea rows={2} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
            <div style={{ display: "flex", gap: 8 }}>
              <Btn small onClick={save}>Spremi</Btn>
              <Btn small kind="ghost" onClick={() => setEdit(false)}>Odustani</Btn>
            </div>
          </>
        )}
      </Card>

      <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA" }}>
        <div style={{ fontWeight: 700, color: S.blue, marginBottom: 8 }}>🏨 Objekt: {objName(worker.objectId) || "Bez objekta"}</div>
        <div style={{ display: "flex", gap: 8, alignItems: "flex-end", flexWrap: "wrap" }}>
          <div style={{ flex: "1 1 150px" }}><Field label="Prebaci na"><ObjectSelect data={data} api={api} value={moveObj} onChange={setMoveObj} /></Field></div>
          <div style={{ flex: "1 1 120px" }}><Field label="Od datuma"><input type="date" value={moveDate} onChange={(e) => setMoveDate(e.target.value)} /></Field></div>
          <div style={{ marginBottom: 10 }}><Btn small onClick={transfer}>Prebaci</Btn></div>
        </div>
        {myAssignments.length > 0 && (
          <div>
            <div style={{ fontSize: 12.5, fontWeight: 700, color: S.blue, marginBottom: 4 }}>Povijest premještaja</div>
            {myAssignments.map((a, i) => (
              <div key={a.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", borderBottom: `1px solid #CBDCEA`, fontSize: 13.5 }}>
                <span className="num" style={{ flex: 1 }}>
                  od {fmtDate(a.from)} · <b>{objName(a.objectId) || "Bez objekta"}</b>
                  {i === 0 ? <span style={{ color: S.green, fontWeight: 700 }}> (trenutno)</span> : ""}
                </span>
                <button onClick={() => api.delAssignment(a, worker.name)} style={{ background: "none", border: "none", color: S.red, fontSize: 15, cursor: "pointer", padding: 4 }}>✕</button>
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card style={{ background: S.greenSoft, borderColor: "#C5DED2" }}>
        <div style={{ fontWeight: 700, color: S.green, marginBottom: 8 }}>
          💶 Satnica: {rateNow(data, worker) > 0 ? money(rateNow(data, worker), wCur(worker)) + "/h" : "nije unesena — dodaj je ovdje"}
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "flex-end", flexWrap: "wrap" }}>
          <div style={{ flex: "1 1 120px" }}><Field label={`Nova satnica (${wCur(worker) === "CZK" ? "Kč" : "€"}/h)`}><input inputMode="decimal" value={rateNew} onChange={(e) => setRateNew(e.target.value)} onKeyDown={(e) => e.key === "Enter" && changeRate()} placeholder={wCur(worker) === "CZK" ? "npr. 160" : "npr. 7.00"} /></Field></div>
          <div style={{ flex: "1 1 130px" }}><Field label="Vrijedi od"><input type="date" value={rateFrom} onChange={(e) => setRateFrom(e.target.value)} /></Field></div>
          <div style={{ marginBottom: 10 }}><Btn small onClick={changeRate}>Promijeni</Btn></div>
        </div>
        <div style={{ fontSize: 12.5, fontWeight: 700, color: S.green, marginBottom: 2 }}>Povijest satnice</div>
        <div className="num" style={{ fontSize: 13.5, padding: "4px 0", color: S.sub }}>početna · <b>{worker.rate > 0 ? money(worker.rate, wCur(worker)) + "/h" : "—"}</b></div>
        {myRates.map((r) => (
          <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "4px 0", borderTop: `1px solid #C5DED2`, fontSize: 13.5 }}>
            <span className="num" style={{ flex: 1 }}>od {fmtDate(r.from)} · <b>{money(r.rate, wCur(worker))}/h</b></span>
            <button onClick={() => api.delRateChange(r, worker.name)} style={{ background: "none", border: "none", color: S.red, fontSize: 15, cursor: "pointer", padding: 4 }}>✕</button>
          </div>
        ))}
      </Card>

      <Card>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <div style={{ fontWeight: 700 }}>Dokumenti i ugovori</div>
          <Btn small kind="ghost" onClick={() => fileRef.current?.click()}>+ Slika ili datoteka</Btn>
          <input ref={fileRef} type="file" accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.txt" multiple onChange={uploadDoc} style={{ display: "none" }} />
        </div>
        {docErr && <div style={{ color: S.red, fontSize: 13, marginBottom: 8 }}>{docErr}</div>}
        {docs === null ? <div style={{ color: S.sub, fontSize: 13.5 }}>Učitavam…</div>
          : docs.length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>Slikaj ugovor ili dodaj PDF.</div>
          : (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8 }}>
              {docs.map((f) => (
                <div key={f.name} onClick={() => openDoc(f)} style={{ aspectRatio: "1", borderRadius: 10, border: `1px solid ${S.line}`, background: "#FAFBF9",
                  display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 6, padding: 8, cursor: "pointer" }}>
                  <span style={{ fontSize: 26 }}>{/\.(jpe?g|png|gif|webp)$/i.test(f.name) ? "🖼" : "📄"}</span>
                  <span style={{ fontSize: 10.5, fontWeight: 600, color: S.sub, textAlign: "center", wordBreak: "break-word", overflow: "hidden" }}>
                    {f.name.replace(/^\d+_/, "").slice(0, 40)}
                  </span>
                </div>
              ))}
            </div>
          )}
      </Card>

      <Card>
        <div style={{ fontWeight: 700, marginBottom: 8 }}>Zadnji upisi sati</div>
        {logs.length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>Još nema upisanih sati.</div>
          : logs.map((l) => (
            <div key={l.id} style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", borderBottom: `1px solid ${S.line}`, fontSize: 14 }}>
              <span>{fmtDate(l.date)} · {logSpan(l)}{objName(l.objectId) ? " · " + objName(l.objectId) : ""}</span>
              <span className="num" style={{ fontWeight: 700 }}>{fmtH(l.hours)}</span>
            </div>
          ))}
      </Card>

      {viewer && (
        <div onClick={() => setViewer(null)} style={{ position: "fixed", inset: 0, background: "rgba(10,14,12,.92)", zIndex: 50,
          display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 16 }}>
          {viewer.image
            ? <img src={viewer.url} alt="" style={{ maxWidth: "100%", maxHeight: "72vh", borderRadius: 12 }} />
            : (
              <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", borderRadius: 14, padding: 22, textAlign: "center", maxWidth: 320 }}>
                <div style={{ fontSize: 42 }}>📄</div>
                <div style={{ fontWeight: 700, marginTop: 8, wordBreak: "break-word" }}>{viewer.name.replace(/^\d+_/, "")}</div>
              </div>
            )}
          <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
            <a href={viewer.url} target="_blank" rel="noreferrer" style={{ background: S.green, color: "#fff", borderRadius: 10, padding: "9px 14px", fontWeight: 700, textDecoration: "none", fontSize: 13.5 }}>⬇ Otvori / preuzmi</a>
            {api.admin && <Btn small kind="ghost" onClick={(e) => { e.stopPropagation(); delDoc({ name: viewer.name }); }}>Obriši</Btn>}
            <Btn small onClick={() => setViewer(null)}>Zatvori</Btn>
          </div>
        </div>
      )}
    </>
  );
}

/* ================================================================== */
/*  OBJEKT — BRZI UNOS + kopiranje dana                                */
/* ================================================================== */
function ObjectDetail({ object, data, api, onBack }) {
  const [mode, setMode] = useState("day");
  const [date, setDate] = useState(todayISO());
  const [month, setMonth] = useState(curMonth());
  const [inputs, setInputs] = useState({});
  const [editLog, setEditLog] = useState(null);
  const [rateEdit, setRateEdit] = useState(String(object.billRate || ""));
  const [billCur, setBillCur] = useState(object.billCur || "EUR");
  const [invoices, setInvoices] = useState(null);
  const [invErr, setInvErr] = useState("");
  const [invViewer, setInvViewer] = useState(null);
  const invFileRef = useRef(null);
  const [showGen, setShowGen] = useState(false);
  const [genMonth, setGenMonth] = useState(curMonth());
  const [genNumber, setGenNumber] = useState("");
  const [genHours, setGenHours] = useState("");
  const [genRate, setGenRate] = useState(String(object.billRate || ""));
  const [genCur, setGenCur] = useState(object.billCur || "EUR");
  const [autoAmount, setAutoAmount] = useState(true);
  const [genAmountManual, setGenAmountManual] = useState("");
  const [genNote, setGenNote] = useState("");

  useEffect(() => { setRateEdit(String(object.billRate || "")); setBillCur(object.billCur || "EUR"); }, [object.id, object.billRate, object.billCur]);

  const loadInvoices = useCallback(async () => {
    if (!api.admin) return;
    const { data: files, error } = await supabase.storage.from("invoices").list(object.id, { limit: 200, sortBy: { column: "created_at", order: "desc" } });
    if (error) { setInvErr(error.message); return; }
    setInvoices(files || []);
  }, [object.id, api.admin]);
  useEffect(() => { loadInvoices(); }, [loadInvoices]);

  const uploadInvoice = async (e) => {
    const files = Array.from(e.target.files || []);
    setInvErr("");
    for (const f of files) {
      if (f.size > 15 * 1024 * 1024) { setInvErr(`"${f.name}" je veća od 15 MB.`); continue; }
      const { error } = await supabase.storage.from("invoices").upload(`${object.id}/${Date.now()}_${f.name}`, f);
      if (error) setInvErr(error.message);
    }
    await api.logDoc(`Dodao fakturu za objekt: ${object.name}`);
    await loadInvoices();
    e.target.value = "";
  };
  const openInvoice = async (f) => {
    const { data: s } = await supabase.storage.from("invoices").createSignedUrl(`${object.id}/${f.name}`, 300);
    if (s?.signedUrl) setInvViewer({ name: f.name, url: s.signedUrl, image: /\.(jpe?g|png|gif|webp)$/i.test(f.name) });
  };
  const delInvoice = async (f) => {
    const { error } = await supabase.storage.from("invoices").remove([`${object.id}/${f.name}`]);
    if (error) { setInvErr(error.message); return; }
    await api.logDoc(`Obrisao fakturu objekta: ${object.name}`);
    setInvViewer(null);
    await loadInvoices();
  };

  const objHoursFor = (mo) => round2(data.logs.filter((l) => l.objectId === object.id && monthKey(l.date) === mo).reduce((s, l) => s + l.hours, 0));
  const openGenerator = () => {
    const hrs = objHoursFor(genMonth);
    setGenHours(String(hrs));
    setGenRate(String(object.billRate || ""));
    setGenCur(object.billCur || "EUR");
    const code = object.name.replace(/[^a-zA-ZčćžšđČĆŽŠĐ]+/g, "").slice(0, 3).toUpperCase() || "OBJ";
    setGenNumber(`${code}-${genMonth}`);
    setAutoAmount(true);
    setGenAmountManual("");
    setGenNote("");
    setShowGen(true);
  };
  const pickGenMonth = (mo) => {
    setGenMonth(mo);
    setGenHours(String(objHoursFor(mo)));
  };
  const genAuto = round2((parseNum(genHours) || 0) * (parseNum(genRate) || 0));
  const submitInvoice = async () => {
    const hours = round2(parseNum(genHours) || 0);
    const rate = round2(parseNum(genRate) || 0);
    const amount = autoAmount ? genAuto : round2(parseNum(genAmountManual) || 0);
    if (!genNumber.trim() || amount <= 0) return;
    const inv = { objectId: object.id, number: genNumber.trim(), period: genMonth, issueDate: todayISO(), hours, rate, amount, currency: genCur, note: genNote };
    if (await api.addObjectInvoice(inv, object.name)) {
      printGeneratedInvoice(inv, object, data.settings || {});
      setShowGen(false);
    }
  };
  const myInvoices = (data.objectInvoices || []).filter((i) => i.objectId === object.id);

  const setIn = (id, patch) => setInputs((p) => ({ ...p, [id]: { ...(p[id] || { from: "", to: "", hours: "" }), ...patch } }));
  const [my, mm] = month.split("-").map(Number);
  const prevM = () => setMonth(mm === 1 ? `${my - 1}-12` : `${my}-${String(mm - 1).padStart(2, "0")}`);
  const nextM = () => setMonth(mm === 12 ? `${my + 1}-01` : `${my}-${String(mm + 1).padStart(2, "0")}`);

  const workers = data.workers.filter((w) => !w.archived).sort((a, b) => {
    const am = a.objectId === object.id ? 0 : 1, bm = b.objectId === object.id ? 0 : 1;
    return am - bm || a.name.localeCompare(b.name);
  });

  const commitDay = (w) => {
    const v = inputs[w.id] || {};
    let hours = parseNum(v.hours), from = v.from || "", to = v.to || "";
    if ((!hours || hours <= 0) && from && to) hours = hoursBetween(from, to);
    if (!hours || hours <= 0) return;
    if (!(from && to)) { from = ""; to = ""; }
    api.addLog({ workerId: w.id, objectId: object.id, date, from, to, hours: round2(hours) }, w.name, object.name);
    setIn(w.id, { from: "", to: "", hours: "" });
  };
  const commitMonth = (w) => {
    const hours = parseNum((inputs[w.id] || {}).hours);
    if (!hours || hours <= 0) return;
    api.addLog({ workerId: w.id, objectId: object.id, date: month + "-01", from: "", to: "", hours: round2(hours), monthly: true }, w.name, object.name);
    setIn(w.id, { from: "", to: "", hours: "" });
  };

  const entries = data.logs.filter((l) => l.objectId === object.id && (mode === "day" ? l.date === date : monthKey(l.date) === month))
    .sort((a, b) => a.date.localeCompare(b.date));
  const wName = (id) => data.workers.find((x) => x.id === id)?.name || "Obrisan radnik";
  const monthHours = (wid) => round2(data.logs.filter((l) => l.workerId === wid && l.objectId === object.id && monthKey(l.date) === (mode === "day" ? monthKey(date) : month)).reduce((s, l) => s + l.hours, 0));

  return (
    <>
      <button onClick={onBack} style={{ background: "none", border: "none", color: S.green, fontWeight: 700, fontSize: 14.5, padding: "2px 0 12px", cursor: "pointer" }}>← Natrag</button>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
        <div style={{ fontSize: 19, fontWeight: 800 }}>🏨 {object.name}</div>
        <div style={{ display: "flex", gap: 6, marginLeft: "auto" }}>
          {["HR", "CZ"].map((c) => (
            <button key={c} onClick={() => (object.country || "HR") !== c && api.setObjectCountry(object, c)} style={{
              padding: "5px 11px", borderRadius: 999, fontWeight: 700, fontSize: 12.5, cursor: "pointer",
              background: (object.country || "HR") === c ? S.blue : "#fff", color: (object.country || "HR") === c ? "#fff" : S.sub,
              border: `1px solid ${(object.country || "HR") === c ? S.blue : S.line}` }}>
              {FLAG[c]} {c}
            </button>
          ))}
        </div>
      </div>

      {api.admin && (() => {
        const mk = mode === "day" ? monthKey(date) : month;
        const objLogs = data.logs.filter((l) => l.objectId === object.id && monthKey(l.date) === mk);
        const hrs = round2(objLogs.reduce((s, l) => s + l.hours, 0));
        const revenue = round2(hrs * (object.billRate || 0));
        const bc = object.billCur || "EUR";
        let costE = 0, costK = 0;
        objLogs.forEach((l) => {
          const wk = data.workers.find((x) => x.id === l.workerId);
          const v = l.hours * rateFor(data, wk, l.date);
          if (wCur(wk) === "CZK") costK += v; else costE += v;
        });
        costE = round2(costE); costK = round2(costK);
        const oPays = data.payments.filter((pp) => !pp.workerId && pp.objectId === object.id && monthKey(pp.date) === mk);
        const extraCost = round2(oPays.filter((pp) => (pp.currency || "EUR") === "EUR").reduce((s, pp) => s + pp.amount, 0));
        const extraCostKc = round2(oPays.filter((pp) => pp.currency === "CZK").reduce((s, pp) => s + pp.amount, 0));
        return (
          <Card style={{ background: S.amberSoft, borderColor: "#EBD9B4" }}>
            <div style={{ fontWeight: 700, color: S.amber, marginBottom: 8 }}>👑 Admin — naplata i dobit</div>
            <Field label={`Objekt plaća (${bc === "CZK" ? "Kč" : "€"} / sat)`}>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input inputMode="decimal" value={rateEdit} onChange={(e) => setRateEdit(e.target.value)} placeholder={bc === "CZK" ? "npr. 250" : "npr. 10.00"} />
                <CurChips small value={billCur} onChange={setBillCur} />
                <Btn small onClick={() => api.setBillRate(object, parseNum(rateEdit) || 0, billCur)}>Spremi</Btn>
              </div>
            </Field>
            {(() => {
              const revE = bc === "EUR" ? revenue : 0;
              const revK = bc === "CZK" ? revenue : 0;
              const profitE = round2(revE - costE - extraCost);
              const profitK = round2(revK - costK - extraCostKc);
              return (
                <div className="num" style={{ fontSize: 14 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>Naplata ({fmtH(hrs)} × {money(object.billRate || 0, bc)})</span><b>{money(revenue, bc)}</b></div>
                  {costE > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>Plaća radnika (€)</span><b>−{eur(costE)}</b></div>}
                  {costK > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>Plaća radnika (Kč)</span><b>−{czk(costK)}</b></div>}
                  {extraCost > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>Troškovi objekta (€)</span><b>−{eur(extraCost)}</b></div>}
                  {extraCostKc > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>Troškovi objekta (Kč)</span><b>−{czk(extraCostKc)}</b></div>}
                  <div style={{ borderTop: `1px dashed #EBD9B4`, paddingTop: 5, marginTop: 2 }}>
                    {(profitE !== 0 || profitK === 0) && (
                      <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 800 }}>
                        <span>Dobit ovaj mjesec (€)</span><span style={{ color: profitE >= 0 ? S.green : S.red }}>{eur(profitE)}</span>
                      </div>
                    )}
                    {profitK !== 0 && (
                      <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 800 }}>
                        <span>Dobit ovaj mjesec (Kč)</span><span style={{ color: profitK >= 0 ? S.green : S.red }}>{czk(profitK)}</span>
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}
            <div style={{ borderTop: `1px dashed #EBD9B4`, marginTop: 10, paddingTop: 10 }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, color: S.amber, marginBottom: 6 }}>Vidljiv zaposlenicima</div>
              {(data.profiles || []).filter((p) => p.role !== "admin").length === 0 ? (
                <div style={{ fontSize: 12.5, color: S.sub }}>Još nema zaposlenika.</div>
              ) : (data.profiles || []).filter((p) => p.role !== "admin").map((p) => {
                const on = (data.objectMembers || []).some((m) => m.object_id === object.id && m.user_id === p.id);
                return (
                  <label key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", fontSize: 13.5, cursor: "pointer" }}>
                    <input type="checkbox" checked={on} onChange={(e) => api.toggleMember(object, p, e.target.checked)} style={{ width: 18, height: 18 }} />
                    {p.name || "(bez imena)"}
                  </label>
                );
              })}
            </div>
          </Card>
        );
      })()}

      {(() => {
        const mk2 = mode === "day" ? monthKey(date) : month;
        const wStats = new Map();
        data.logs.filter((l) => l.objectId === object.id && monthKey(l.date) === mk2).forEach((l) => {
          wStats.set(l.workerId, round2((wStats.get(l.workerId) || 0) + l.hours));
        });
        data.workers.filter((w) => !w.archived && w.objectId === object.id).forEach((w) => {
          if (!wStats.has(w.id)) wStats.set(w.id, 0);
        });
        const rows2 = [...wStats.entries()]
          .map(([wid, h]) => ({ w: data.workers.find((x) => x.id === wid), h }))
          .filter((r) => r.w)
          .sort((a, b) => a.w.name.localeCompare(b.w.name, "hr"));
        if (!rows2.length) return null;
        const tot = round2(rows2.reduce((s, r) => s + r.h, 0));
        return (
          <Card>
            <div style={{ fontWeight: 700, marginBottom: 6 }}>👷 Radnici na objektu ovaj mjesec</div>
            {rows2.map((r) => (
              <div key={r.w.id} className="num" style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13.5 }}>
                <span style={{ fontWeight: 600 }}>{r.w.name}{r.w.objectId === object.id ? <span style={{ color: S.blue, fontSize: 11, fontWeight: 700 }}> ★</span> : ""}</span>
                <span style={{ fontWeight: 700 }}>{fmtH(r.h)}</span>
              </div>
            ))}
            <div className="num" style={{ display: "flex", justifyContent: "space-between", paddingTop: 7, fontWeight: 800 }}>
              <span>Ukupno</span><span>{fmtH(tot)}</span>
            </div>
          </Card>
        );
      })()}

      {api.admin && (
        <Card>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
            <div style={{ fontWeight: 700 }}>🧾 Fakture i dokumenti objekta</div>
            <div style={{ display: "flex", gap: 6 }}>
              <Btn small onClick={() => (showGen ? setShowGen(false) : openGenerator())}>{showGen ? "Zatvori" : "⚡ Generiraj fakturu"}</Btn>
              <Btn small kind="ghost" onClick={() => invFileRef.current?.click()}>+ Dodaj datoteku</Btn>
            </div>
            <input ref={invFileRef} type="file" accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.txt" multiple onChange={uploadInvoice} style={{ display: "none" }} />
          </div>

          {showGen && (
            <div style={{ background: S.greenSoft, borderRadius: 10, padding: "12px 14px", marginBottom: 14 }}>
              <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 10 }}>
                Podaci su povučeni iz upisanih sati i naplate objekta — po potrebi ih slobodno promijeni prije nego stisneš Generiraj.
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <div style={{ flex: 1 }}><Field label="Mjesec"><input type="month" value={genMonth} onChange={(e) => pickGenMonth(e.target.value)} /></Field></div>
                <div style={{ flex: 1 }}><Field label="Broj računa"><input value={genNumber} onChange={(e) => setGenNumber(e.target.value)} placeholder="npr. ADR-2026-08" /></Field></div>
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <div style={{ flex: 1 }}><Field label="Sati"><input inputMode="decimal" value={genHours} onChange={(e) => setGenHours(e.target.value)} /></Field></div>
                <div style={{ flex: 1 }}>
                  <Field label={`Cijena po satu (${genCur === "CZK" ? "Kč" : "€"})`}>
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <input inputMode="decimal" value={genRate} onChange={(e) => setGenRate(e.target.value)} />
                      <CurChips small value={genCur} onChange={setGenCur} />
                    </div>
                  </Field>
                </div>
              </div>
              <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, fontSize: 13.5, cursor: "pointer" }}>
                <input type="checkbox" checked={autoAmount} onChange={(e) => { setAutoAmount(e.target.checked); if (!e.target.checked) setGenAmountManual(String(genAuto)); }} style={{ width: 18, height: 18 }} />
                Iznos automatski (sati × cijena)
              </label>
              <Field label={`Iznos za naplatu (${genCur === "CZK" ? "Kč" : "€"})`}>
                {autoAmount
                  ? <div style={{ background: "#fff", border: `1px solid ${S.line}`, borderRadius: 10, padding: "11px 12px", fontWeight: 700 }}>{money(genAuto, genCur)}</div>
                  : <input inputMode="decimal" value={genAmountManual} onChange={(e) => setGenAmountManual(e.target.value)} />}
              </Field>
              <Field label="Napomena na fakturi (nije obavezno)"><input value={genNote} onChange={(e) => setGenNote(e.target.value)} placeholder="npr. uključuje prekovremene" /></Field>
              <Btn onClick={submitInvoice} style={{ width: "100%" }}>🧾 Generiraj i spremi fakturu</Btn>
            </div>
          )}

          {invErr && <div style={{ color: S.red, fontSize: 13, marginBottom: 8 }}>{invErr}</div>}

          {myInvoices.length > 0 && (
            <div style={{ marginBottom: 14 }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, color: S.sub, marginBottom: 6 }}>Generirane fakture</div>
              {myInvoices.map((inv) => (
                <div key={inv.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderBottom: `1px solid ${S.line}` }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 600, fontSize: 13.5 }}>{inv.number}</div>
                    <div className="num" style={{ fontSize: 12, color: S.sub }}>{fmtDate(inv.issueDate)} · {fmtH(inv.hours)} · {money(inv.amount, inv.currency)}</div>
                  </div>
                  <Btn small kind="ghost" onClick={() => printGeneratedInvoice(inv, object, data.settings || {})}>🖨</Btn>
                  <button onClick={() => api.delObjectInvoice(inv, object.name)} style={{ background: "none", border: "none", color: S.red, fontSize: 16, cursor: "pointer", padding: 4 }}>✕</button>
                </div>
              ))}
            </div>
          )}

          {invoices === null ? <div style={{ color: S.sub, fontSize: 13.5 }}>Učitavam…</div>
            : invoices.length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>Slikaj ili dodaj PDF fakture koje stižu od hotela — ostat će spremljene ovdje.</div>
            : (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8 }}>
                {invoices.map((f) => (
                  <div key={f.name} onClick={() => openInvoice(f)} style={{ aspectRatio: "1", borderRadius: 10, border: `1px solid ${S.line}`, background: "#FAFBF9",
                    display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 6, padding: 8, cursor: "pointer" }}>
                    <span style={{ fontSize: 26 }}>{/\.(jpe?g|png|gif|webp)$/i.test(f.name) ? "🖼" : "📄"}</span>
                    <span style={{ fontSize: 10.5, fontWeight: 600, color: S.sub, textAlign: "center", wordBreak: "break-word", overflow: "hidden" }}>
                      {f.name.replace(/^\d+_/, "").slice(0, 40)}
                    </span>
                  </div>
                ))}
              </div>
            )}
        </Card>
      )}

      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
        {[["day", "Unos po danu"], ["month", "Sati za cijeli mjesec"]].map(([id, label]) => (
          <button key={id} onClick={() => setMode(id)} style={{ flex: 1, padding: "10px 8px", borderRadius: 10, fontWeight: 700, fontSize: 13.5, cursor: "pointer",
            background: mode === id ? S.blue : "#fff", color: mode === id ? "#fff" : S.sub, border: `1px solid ${mode === id ? S.blue : S.line}` }}>{label}</button>
        ))}
      </div>

      {mode === "day" ? (
        <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA", padding: "10px 14px" }}>
          <Field label="Datum"><input type="date" value={date} onChange={(e) => setDate(e.target.value)} /></Field>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
            <DayBadge iso={date} />
            <Btn small kind="ghost" onClick={() => api.copyDay(object, addDaysISO(date, -1), date)} style={{ marginLeft: "auto" }}>
              📋 Kopiraj jučerašnji dan
            </Btn>
          </div>
        </Card>
      ) : (
        <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA", padding: "10px 14px" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <Btn small kind="ghost" onClick={prevM}>←</Btn>
            <div style={{ fontWeight: 800 }}>{MONTHS[mm - 1]} {my}.</div>
            <Btn small kind="ghost" onClick={nextM}>→</Btn>
          </div>
        </Card>
      )}

      {workers.map((w) => {
        const v = inputs[w.id] || { from: "", to: "", hours: "" };
        return (
          <Card key={w.id} style={{ padding: "12px 14px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div>
                <span style={{ fontWeight: 700, fontSize: 15 }}>{w.name}</span>
                {w.objectId === object.id && <span style={{ marginLeft: 6, fontSize: 11, color: S.blue, fontWeight: 700 }}>★ ovaj objekt</span>}
              </div>
              <div className="num" style={{ fontSize: 12.5, color: S.sub }}>{fmtH(monthHours(w.id))} ovdje ovaj mj.</div>
            </div>
            {mode === "day" ? (
              <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <input type="time" value={v.from} onChange={(e) => setIn(w.id, { from: e.target.value })} style={{ flex: 1.2, padding: "9px 6px" }} />
                <span style={{ color: S.sub }}>–</span>
                <input type="time" value={v.to} onChange={(e) => setIn(w.id, { to: e.target.value })} style={{ flex: 1.2, padding: "9px 6px" }} />
                <input inputMode="decimal" placeholder="ili sati" value={v.hours} onChange={(e) => setIn(w.id, { hours: e.target.value })} style={{ flex: 1, padding: "9px 8px" }} />
                <Btn small onClick={() => commitDay(w)}>Upiši</Btn>
              </div>
            ) : (
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input inputMode="decimal" placeholder="ukupno sati za mjesec" value={v.hours} onChange={(e) => setIn(w.id, { hours: e.target.value })} />
                <Btn small onClick={() => commitMonth(w)}>Upiši</Btn>
              </div>
            )}
          </Card>
        );
      })}

      {entries.length > 0 && (
        <Card>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>Upisi — {mode === "day" ? fmtDate(date) : `${MONTHS[mm - 1]} ${my}.`}</div>
          {entries.map((l) => (
            <div key={l.id} style={{ borderBottom: `1px solid ${S.line}` }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0" }}>
                <div style={{ flex: 1 }}>
                  <span style={{ fontWeight: 600, fontSize: 14 }}>{wName(l.workerId)}</span>
                  <div className="num" style={{ fontSize: 12.5, color: S.sub }}>{mode === "month" ? fmtDate(l.date) + " · " : ""}{logSpan(l)}</div>
                </div>
                <div className="num" style={{ fontWeight: 700 }}>{fmtH(l.hours)}</div>
                {(api.admin || l.createdBy === api.uid()) ? (
                  <>
                    <button onClick={() => setEditLog(editLog === l.id ? null : l.id)} style={{ background: "none", border: "none", color: S.blue, fontSize: 15, cursor: "pointer", padding: 4 }}>✎</button>
                    <button onClick={() => api.delLog(l, wName(l.workerId))} style={{ background: "none", border: "none", color: S.red, fontSize: 16, cursor: "pointer", padding: 4 }}>✕</button>
                  </>
                ) : <span title="Upisao kolega" style={{ fontSize: 13, color: S.sub, padding: 4 }}>👤</span>}
              </div>
              {editLog === l.id && (
                <LogEditor log={l} onCancel={() => setEditLog(null)}
                  onSave={async (patch) => { if (await api.updLog(l, patch, wName(l.workerId))) setEditLog(null); }} />
              )}
            </div>
          ))}
          <div className="num" style={{ display: "flex", justifyContent: "space-between", fontWeight: 800, paddingTop: 8 }}>
            <span>Ukupno</span><span>{fmtH(round2(entries.reduce((s, l) => s + l.hours, 0)))}</span>
          </div>
        </Card>
      )}

      {invViewer && (
        <div onClick={() => setInvViewer(null)} style={{ position: "fixed", inset: 0, background: "rgba(10,14,12,.92)", zIndex: 50,
          display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 16 }}>
          {invViewer.image
            ? <img src={invViewer.url} alt="" style={{ maxWidth: "100%", maxHeight: "72vh", borderRadius: 12 }} />
            : (
              <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", borderRadius: 14, padding: 22, textAlign: "center", maxWidth: 320 }}>
                <div style={{ fontSize: 42 }}>📄</div>
                <div style={{ fontWeight: 700, marginTop: 8, wordBreak: "break-word" }}>{invViewer.name.replace(/^\d+_/, "")}</div>
              </div>
            )}
          <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
            <a href={invViewer.url} target="_blank" rel="noreferrer" style={{ background: S.green, color: "#fff", borderRadius: 10, padding: "9px 14px", fontWeight: 700, textDecoration: "none", fontSize: 13.5 }}>⬇ Otvori / preuzmi</a>
            <Btn small kind="ghost" onClick={(e) => { e.stopPropagation(); delInvoice({ name: invViewer.name }); }}>Obriši</Btn>
            <Btn small onClick={() => setInvViewer(null)}>Zatvori</Btn>
          </div>
        </div>
      )}
    </>
  );
}

/* ================================================================== */
/*  SATI + podsjetnik na neupisane dane                                */
/* ================================================================== */
function HoursTab({ data, api }) {
  const [form, setForm] = useState({ workerId: "", objectId: "", date: todayISO(), from: "07:00", to: "15:00", note: "" });
  const h = hoursBetween(form.from, form.to);
  const w = data.workers.find((x) => x.id === form.workerId);
  const objName = (id) => data.objects.find((o) => o.id === id)?.name || "";

  /* radni dani ovog mjeseca (do danas) bez ijednog upisa */
  const mk = curMonth();
  const missing = [];
  if (data.logs.some((l) => monthKey(l.date) === mk)) {
    for (let d = mk + "-01"; d < todayISO(); d = addDaysISO(d, 1)) {
      if (!isWeekend(d) && !holidayName(d) && !data.logs.some((l) => l.date === d)) missing.push(d);
    }
  }

  const pickWorker = (id) => {
    const wk = data.workers.find((x) => x.id === id);
    setForm({ ...form, workerId: id, objectId: wk?.objectId || form.objectId });
  };
  const workerOptions = workersAtObject(data, form.objectId);
  const pickObject = (id) => {
    const stillValid = workersAtObject(data, id).some((x) => x.id === form.workerId);
    setForm({ ...form, objectId: id, workerId: stillValid ? form.workerId : "" });
  };
  const add = () => {
    if (!form.workerId || !form.date || h <= 0) return;
    api.addLog({ workerId: form.workerId, objectId: form.objectId, date: form.date, from: form.from, to: form.to, hours: h, note: form.note }, w?.name || "", objName(form.objectId));
    setForm({ ...form, note: "" });
  };
  const recent = [...data.logs].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 15);
  const wName = (id) => data.workers.find((x) => x.id === id)?.name || "Obrisan radnik";
  const [editLog, setEditLog] = useState(null);

  return (
    <>
      <div style={{ fontSize: 17, fontWeight: 700, marginBottom: 6 }}>{api.t("hoursTitle")}</div>

      {missing.length > 0 && (
        <Card style={{ background: S.amberSoft, borderColor: "#EBD9B4", padding: "10px 14px" }}>
          <div style={{ fontSize: 13.5 }}>
            <b style={{ color: S.amber }}>⏰ Dani bez ijednog upisa ovaj mjesec:</b>{" "}
            <span className="num">{missing.slice(0, 8).map((d) => fmtDate(d).slice(0, -1)).join(", ")}{missing.length > 8 ? ` +${missing.length - 8}` : ""}</span>
          </div>
        </Card>
      )}

      {data.workers.filter((x) => !x.archived).length === 0 ? <Empty text="Prvo dodaj radnika." /> : (
        <Card>
          <Field label={api.t("object")}><ObjectSelect data={data} api={api} value={form.objectId} onChange={pickObject} /></Field>
          <Field label={form.objectId ? `${api.t("worker")} (${api.t("object").toLowerCase()})` : api.t("worker")}>
            <select value={form.workerId} onChange={(e) => pickWorker(e.target.value)}>
              <option value="">{api.t("choose")}</option>
              {workerOptions.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
            </select>
            {form.objectId && workerOptions.length === 0 && (
              <div style={{ fontSize: 12, color: S.sub, marginTop: 4 }}>Nitko još nije upisan na ovom objektu.</div>
            )}
          </Field>
          <Field label={api.t("date")}><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></Field>
          <div style={{ marginBottom: 8 }}><DayBadge iso={form.date} /></div>
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}><Field label={api.t("from")}><input type="time" value={form.from} onChange={(e) => setForm({ ...form, from: e.target.value })} /></Field></div>
            <div style={{ flex: 1 }}><Field label={api.t("to")}><input type="time" value={form.to} onChange={(e) => setForm({ ...form, to: e.target.value })} /></Field></div>
          </div>
          <Field label="Napomena (opcionalno)"><input value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
          <div style={{ background: S.greenSoft, borderRadius: 10, padding: "10px 12px", marginBottom: 12, display: "flex", justifyContent: "space-between", fontWeight: 700 }}>
            <span>{api.t("total")}: <span className="num">{fmtH(h)}</span></span>
            {w && <span className="num" style={{ color: S.green }}>{money(round2(h * rateFor(data, w, form.date)), wCur(w))}</span>}
          </div>
          <Btn onClick={add} style={{ width: "100%" }}>{api.t("addHours")}</Btn>
        </Card>
      )}
      {recent.length > 0 && (
        <Card>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{api.t("recent")}</div>
          {recent.map((l) => (
            <div key={l.id} style={{ borderBottom: `1px solid ${S.line}` }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0" }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 600, fontSize: 14 }}>{wName(l.workerId)}</div>
                  <div style={{ fontSize: 12.5, color: S.sub }}>
                    {fmtDate(l.date)} · {logSpan(l)}{objName(l.objectId) ? " · " + objName(l.objectId) : ""}{l.note && !l.monthly ? " · " + l.note : ""}
                  </div>
                </div>
                <div className="num" style={{ fontWeight: 700 }}>{fmtH(l.hours)}</div>
                {(api.admin || l.createdBy === api.uid()) ? (
                  <>
                    <button onClick={() => setEditLog(editLog === l.id ? null : l.id)} style={{ background: "none", border: "none", color: S.blue, fontSize: 15, cursor: "pointer", padding: 4 }}>✎</button>
                    <button onClick={() => api.delLog(l, wName(l.workerId))} style={{ background: "none", border: "none", color: S.red, fontSize: 16, cursor: "pointer", padding: 4 }}>✕</button>
                  </>
                ) : <span title="Upisao kolega" style={{ fontSize: 13, color: S.sub, padding: 4 }}>👤</span>}
              </div>
              {editLog === l.id && (
                <LogEditor log={l} onCancel={() => setEditLog(null)}
                  onSave={async (patch) => { if (await api.updLog(l, patch, wName(l.workerId))) setEditLog(null); }} />
              )}
            </div>
          ))}
        </Card>
      )}
    </>
  );
}

/* ================================================================== */
/*  AVANSI, BONUSI I TROŠKOVI                                          */
/* ================================================================== */
function PaymentsTab({ data, api }) {
  const [target, setTarget] = useState("radnik"); // radnik | objekt
  const [filterObj, setFilterObj] = useState("");
  const [cur, setCur] = useState("EUR"); // EUR | CZK
  const [form, setForm] = useState({ workerId: "", objectId: "", date: todayISO(), type: "avans", amount: "", note: "", deduct: true });
  const wName = (id) => data.workers.find((x) => x.id === id)?.name || "Obrisan radnik";
  const oName = (id) => data.objects.find((o) => o.id === id)?.name || "Obrisan objekt";
  const add = () => {
    const raw = parseNum(form.amount);
    if (!form.date || !raw) return;
    const m = { amount: round2(raw), currency: cur };
    if (target === "radnik") {
      if (!form.workerId) return;
      api.addPayment({ workerId: form.workerId, objectId: "", date: form.date, type: form.type, ...m, note: form.note,
        deduct: form.type === "avans" ? true : form.type === "bonus" ? false : form.deduct }, wName(form.workerId));
    } else {
      if (!form.objectId) return;
      api.addPayment({ workerId: "", objectId: form.objectId, date: form.date, type: form.type === "bonus" ? "ostalo" : form.type,
        ...m, note: form.note, deduct: false }, "🏨 " + oName(form.objectId));
    }
    setForm({ ...form, amount: "", note: "" });
  };
  const recent = [...data.payments].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 15);
  const tagStyle = (p) => p.type === "avans" ? { color: S.amber, bg: S.amberSoft }
    : p.type === "bonus" ? { color: S.green, bg: S.greenSoft }
    : p.deduct ? { color: S.red, bg: S.redSoft } : { color: S.sub, bg: "#EEF0ED" };
  return (
    <>
      <div style={{ fontSize: 17, fontWeight: 700, marginBottom: 10 }}>{api.t("paymentsTitle")}</div>
      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
        {[["radnik", api.t("forWorker")], ["objekt", api.t("forObject")]].map(([id, label]) => (
          <button key={id} onClick={() => setTarget(id)} style={{ flex: 1, padding: "10px 8px", borderRadius: 10, fontWeight: 700, fontSize: 13.5, cursor: "pointer",
            background: target === id ? S.blue : "#fff", color: target === id ? "#fff" : S.sub, border: `1px solid ${target === id ? S.blue : S.line}` }}>{label}</button>
        ))}
      </div>
      {target === "radnik" && data.workers.filter((x) => !x.archived).length === 0 ? <Empty text="Prvo dodaj radnika." /> : (
        <Card>
          {target === "radnik" ? (
            <>
              <Field label="Filter po objektu (nije obavezno, samo za lakše pronalaženje)">
                <ObjectSelect data={data} api={api} value={filterObj} onChange={(v) => {
                  const stillValid = workersAtObject(data, v).some((x) => x.id === form.workerId);
                  const ob = data.objects.find((o) => o.id === v);
                  setFilterObj(v);
                  setForm({ ...form, workerId: stillValid ? form.workerId : "" });
                  if (v && ob && !stillValid) setCur(countryCur(ob.country));
                }} />
              </Field>
              <Field label={filterObj ? `${api.t("worker")} (${api.t("object").toLowerCase()})` : api.t("worker")}>
                <select value={form.workerId} onChange={(e) => {
                  const wk = data.workers.find((x) => x.id === e.target.value);
                  setForm({ ...form, workerId: e.target.value });
                  if (wk) setCur(wCur(wk));
                }}>
                  <option value="">— odaberi —</option>
                  {workersAtObject(data, filterObj).map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
                </select>
                {filterObj && workersAtObject(data, filterObj).length === 0 && (
                  <div style={{ fontSize: 12, color: S.sub, marginTop: 4 }}>Nitko još nije upisan na ovom objektu.</div>
                )}
              </Field>
              <Field label={api.t("type")}>
                <select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
                  <option value="avans">Avans (odbija se od plaće)</option>
                  <option value="bonus">Bonus (dodaje se na plaću)</option>
                  <option value="gorivo">Gorivo</option>
                  <option value="ostalo">Ostali trošak</option>
                </select>
              </Field>
            </>
          ) : (
            <>
              <Field label={api.t("object")}>
                <ObjectSelect data={data} api={api} value={form.objectId} onChange={(v) => {
                  const ob = data.objects.find((o) => o.id === v);
                  setForm({ ...form, objectId: v });
                  if (ob) setCur(countryCur(ob.country));
                }} />
              </Field>
              <Field label={api.t("type")}>
                <select value={form.type === "bonus" ? "ostalo" : form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
                  <option value="avans">Avans</option>
                  <option value="gorivo">Gorivo</option>
                  <option value="ostalo">Ostali trošak</option>
                </select>
              </Field>
              <div style={{ fontSize: 12.5, color: S.sub, margin: "-4px 0 10px" }}>
                Trošak se veže na objekt (ne odbija se nijednom radniku) i admin ga vidi u dobiti objekta.
              </div>
            </>
          )}
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}><Field label={api.t("date")}><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></Field></div>
            <div style={{ flex: 1 }}>
              <Field label={cur === "CZK" ? "Iznos (Kč)" : api.t("amount")}>
                <input inputMode="decimal" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} placeholder="0.00" />
              </Field>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10 }}>
            {[["EUR", "€ Euro"], ["CZK", "Kč Kruna"]].map(([id, label]) => (
              <button key={id} onClick={() => setCur(id)} style={{ padding: "7px 14px", borderRadius: 999, fontWeight: 700, fontSize: 13, cursor: "pointer",
                background: cur === id ? S.green : "#fff", color: cur === id ? "#fff" : S.sub, border: `1px solid ${cur === id ? S.green : S.line}` }}>{label}</button>
            ))}
          </div>
          <Field label={form.type === "bonus" ? "Za što je bonus" : "Napomena"}>
            <input value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })}
              placeholder={form.type === "bonus" ? "npr. dobro odrađena sezona…" : "npr. gorivo za put"} />
          </Field>
          {target === "radnik" && form.type !== "avans" && form.type !== "bonus" && (
            <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, fontSize: 14 }}>
              <input type="checkbox" checked={form.deduct} onChange={(e) => setForm({ ...form, deduct: e.target.checked })} style={{ width: 18, height: 18 }} />
              Odbij radniku od plaće (inače je trošak firme)
            </label>
          )}
          <Btn onClick={add} style={{ width: "100%" }}>{api.t("save")}</Btn>
        </Card>
      )}
      {recent.length > 0 && (
        <Card>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{api.t("recent")}</div>
          {recent.map((p) => {
            const t = tagStyle(p);
            return (
              <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: `1px solid ${S.line}` }}>
                <Tag color={t.color} bg={t.bg}>{TYPE_LABEL[p.type] || p.type}</Tag>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 600, fontSize: 14 }}>{p.workerId ? wName(p.workerId) : "🏨 " + oName(p.objectId)}</div>
                  <div style={{ fontSize: 12.5, color: S.sub }}>{fmtDate(p.date)}{p.note ? " · " + p.note : ""}</div>
                </div>
                <div className="num" style={{ fontWeight: 700, color: p.type === "bonus" ? S.green : S.ink }}>{p.type === "bonus" ? "+" : ""}{money(p.amount, p.currency)}</div>
                <button onClick={() => api.delPayment(p, p.workerId ? wName(p.workerId) : "🏨 " + oName(p.objectId))} style={{ background: "none", border: "none", color: S.red, fontSize: 16, cursor: "pointer", padding: 4 }}>✕</button>
              </div>
            );
          })}
        </Card>
      )}
    </>
  );
}

/* ================================================================== */
/*  OBRAČUN: mjesec/godina, isplaćeno, PDF, naplata, grafovi           */
/* ================================================================== */
function calcRows(data, filterFn, objectId) {
  return data.workers.map((w) => {
    const logs = data.logs.filter((l) => l.workerId === w.id && filterFn(l.date) && (!objectId || l.objectId === objectId));
    const pays = data.payments.filter((p) => p.workerId === w.id && filterFn(p.date));
    const hours = round2(logs.reduce((s, l) => s + l.hours, 0));
    const grossAll = round2(logs.reduce((s, l) => s + l.hours * rateFor(data, w, l.date), 0));
    const gross = wCur(w) === "CZK" ? 0 : grossAll;
    const grossKc = wCur(w) === "CZK" ? grossAll : 0;
    const rateSet = [...new Set(logs.map((l) => rateFor(data, w, l.date)))];
    const by = (c) => {
      const mine = pays.filter((p) => (p.currency || "EUR") === c);
      return {
        bonuses: round2(mine.filter((p) => p.type === "bonus").reduce((s, p) => s + p.amount, 0)),
        advances: round2(mine.filter((p) => p.type === "avans").reduce((s, p) => s + p.amount, 0)),
        deductions: round2(mine.filter((p) => p.type !== "avans" && p.type !== "bonus" && p.deduct).reduce((s, p) => s + p.amount, 0)),
        firmCosts: round2(mine.filter((p) => p.type !== "avans" && p.type !== "bonus" && !p.deduct).reduce((s, p) => s + p.amount, 0)),
      };
    };
    const e = by("EUR"), k = by("CZK");
    return { w, logs, pays, hours, gross, grossKc, rateSet,
      bonuses: e.bonuses, advances: e.advances, deductions: e.deductions, firmCosts: e.firmCosts,
      net: round2(gross + e.bonuses - e.advances - e.deductions),
      czk: { ...k, gross: grossKc, net: round2(grossKc + k.bonuses - k.advances - k.deductions) } };
  }).filter((r) => (objectId ? r.hours > 0 : (r.hours > 0 || r.pays.length > 0)));
}

function ReportTab({ data, api, admin }) {
  const [view, setView] = useState("month"); // month | year
  const [month, setMonth] = useState(curMonth());
  const [open, setOpen] = useState(null);
  const [uplata, setUplata] = useState({});
  const [objFilter, setObjFilter] = useState(""); // "" = svi objekti
  const [selected, setSelected] = useState(() => new Set());
  const [bulkBusy, setBulkBusy] = useState(false);
  const [y, m] = month.split("-").map(Number);
  const prev = () => setMonth(m === 1 ? `${y - 1}-12` : `${y}-${String(m - 1).padStart(2, "0")}`);
  const next = () => setMonth(m === 12 ? `${y + 1}-01` : `${y}-${String(m + 1).padStart(2, "0")}`);
  const objName = (id) => data.objects.find((o) => o.id === id)?.name || "";
  const objFilterName = objFilter ? objName(objFilter) : "";

  const inMonth = (d) => monthKey(d) === month;
  useEffect(() => { setSelected(new Set()); }, [month, view, objFilter]);
  const inYear = (d) => (d || "").slice(0, 4) === String(y);
  const rows = calcRows(data, view === "month" ? inMonth : inYear, objFilter || undefined).sort((a, b) => a.w.name.localeCompare(b.w.name, "hr"));
  const unpaidRows = view === "month" ? rows.filter((r) => !paidFor(data, r.w.id, month)) : [];
  const allSelected = unpaidRows.length > 0 && unpaidRows.every((r) => selected.has(r.w.id));
  const toggleSelectAll = () => setSelected(allSelected ? new Set() : new Set(unpaidRows.map((r) => r.w.id)));
  const toggleSelect = (id, e) => {
    e.stopPropagation();
    setSelected((prev) => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
  };
  const payMany = async (list) => {
    setBulkBusy(true);
    for (const r of list) { await api.markPaid(r.w, month, r.net, r.czk.net); }
    setSelected(new Set());
    setBulkBusy(false);
  };
  const periodLogs = data.logs.filter((l) => (view === "month" ? inMonth : inYear)(l.date) && (!objFilter || l.objectId === objFilter));
  const periodLabel = view === "month" ? `${MONTHS[m - 1]} ${y}.` : `${y}. godina`;

  const totals = rows.reduce((t, r) => ({
    hours: round2(t.hours + r.hours), gross: round2(t.gross + r.gross),
    net: round2(t.net + r.net), firm: round2(t.firm + r.firmCosts), bonus: round2(t.bonus + r.bonuses),
    netKc: round2(t.netKc + r.czk.net), firmKc: round2(t.firmKc + r.czk.firmCosts),
    grossKc: round2(t.grossKc + r.grossKc), bonusKc: round2(t.bonusKc + r.czk.bonuses),
  }), { hours: 0, gross: 0, net: 0, firm: 0, bonus: 0, netKc: 0, firmKc: 0, grossKc: 0, bonusKc: 0 });

  const objPays = data.payments.filter((p) => !p.workerId && p.objectId && (view === "month" ? inMonth : inYear)(p.date));
  const objCostsTotal = round2(objPays.filter((p) => (p.currency || "EUR") === "EUR").reduce((s, p) => s + p.amount, 0));
  const objCostsTotalKc = round2(objPays.filter((p) => p.currency === "CZK").reduce((s, p) => s + p.amount, 0));
  const byObject = new Map();
  periodLogs.forEach((l) => {
    const key = l.objectId || "__none__";
    const wk = data.workers.find((x) => x.id === l.workerId);
    const ob = data.objects.find((o) => o.id === l.objectId);
    const cur = byObject.get(key) || { hours: 0, gross: 0, grossKc: 0, revenue: 0, revenueKc: 0, costs: 0, costsKc: 0 };
    cur.hours = round2(cur.hours + l.hours);
    const gv = l.hours * rateFor(data, wk, l.date);
    if (wCur(wk) === "CZK") cur.grossKc = round2(cur.grossKc + gv); else cur.gross = round2(cur.gross + gv);
    const rv = l.hours * (ob?.billRate || 0);
    if ((ob?.billCur || "EUR") === "CZK") cur.revenueKc = round2(cur.revenueKc + rv); else cur.revenue = round2(cur.revenue + rv);
    byObject.set(key, cur);
  });
  objPays.forEach((p) => {
    const cur = byObject.get(p.objectId) || { hours: 0, gross: 0, grossKc: 0, revenue: 0, revenueKc: 0, costs: 0, costsKc: 0 };
    if (p.currency === "CZK") cur.costsKc = round2((cur.costsKc || 0) + p.amount);
    else cur.costs = round2((cur.costs || 0) + p.amount);
    byObject.set(p.objectId, cur);
  });
  const totalRevenue = round2([...byObject.values()].reduce((s, v) => s + v.revenue, 0));
  const totalRevenueKc = round2([...byObject.values()].reduce((s, v) => s + (v.revenueKc || 0), 0));
  const totalCost = round2(totals.gross + totals.bonus + totals.firm + objCostsTotal);
  const totalCostKc = round2(totals.grossKc + totals.bonusKc + totals.firmKc + objCostsTotalKc);
  const totalProfitKc = round2(totalRevenueKc - totalCostKc);
  const totalProfit = round2(totalRevenue - totalCost);

  /* grafovi: zadnjih 6 mjeseci */
  const last6 = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date(y, m - 1 - i, 1);
    const mk2 = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    const ls = data.logs.filter((l) => monthKey(l.date) === mk2);
    const hrs = round2(ls.reduce((s, l) => s + l.hours, 0));
    let profit = 0, profitKc = 0;
    if (admin) {
      let revE = 0, revK = 0, costE2 = 0, costK2 = 0;
      ls.forEach((l) => {
        const ob = data.objects.find((o) => o.id === l.objectId);
        const wk = data.workers.find((x) => x.id === l.workerId);
        const rv = l.hours * (ob?.billRate || 0);
        if ((ob?.billCur || "EUR") === "CZK") revK += rv; else revE += rv;
        const cv = l.hours * rateFor(data, wk, l.date);
        if (wCur(wk) === "CZK") costK2 += cv; else costE2 += cv;
      });
      const extraOf = (c) => data.payments.filter((p) => monthKey(p.date) === mk2 && (p.currency || "EUR") === c && (!p.workerId ? !!p.objectId : (p.type === "bonus" || (!p.deduct && p.type !== "avans")))).reduce((s, p) => s + p.amount, 0);
      profit = round2(revE - costE2 - extraOf("EUR"));
      profitKc = round2(revK - costK2 - extraOf("CZK"));
    }
    last6.push({ label: MONTHS_SHORT[d.getMonth()], hours: hrs, profit, profitKc });
  }

  /* PDF obračun za radnika */
  const printPayslip = (r) => {
    const st = api.settings;
    const gc = wCur(r.w);
    const logsHtml = r.logs.sort((a, b) => a.date.localeCompare(b.date)).map((l) =>
      `<tr><td>${fmtDate(l.date)}</td><td>${logSpan(l)}</td><td>${objName(l.objectId) || ""}</td><td class="right">${fmtH(l.hours)}</td><td class="right">${money(round2(l.hours * rateFor(data, r.w, l.date)), gc)}</td></tr>`).join("");
    const paysHtml = r.pays.map((p) =>
      `<tr><td>${fmtDate(p.date)}</td><td colspan="2">${TYPE_LABEL[p.type]}${p.note ? " — " + p.note : ""}</td><td></td><td class="right">${p.type === "bonus" ? "+" : p.deduct ? "−" : ""}${money(p.amount, p.currency)}</td></tr>`).join("");
    printDoc(`Obračun ${r.w.name}`, `
      ${firmHead(st)}
      <h1>Obračun plaće — ${r.w.name}</h1><h2>${periodLabel}</h2>
      <table><tr><th>Datum</th><th>Vrijeme</th><th>Objekt</th><th class="right">Sati</th><th class="right">Iznos</th></tr>
      ${logsHtml}${paysHtml}
      <tr class="tot"><td colspan="3">UKUPNO ${fmtH(r.hours)}</td><td></td><td class="right">${[r.net !== 0 || r.czk.net === 0 ? eur(r.net) : "", r.czk.net !== 0 ? (r.czk.net > 0 ? "" : "−") + czk(Math.abs(r.czk.net)) : ""].filter(Boolean).join("<br>")}</td></tr></table>
      <div class="muted">Zarada ${gc === "CZK" ? czk(r.grossKc) : eur(r.gross)}${r.bonuses ? " + bonus " + eur(r.bonuses) : ""}${r.czk.bonuses ? " + bonus " + czk(r.czk.bonuses) : ""}${r.advances ? " − avans " + eur(r.advances) : ""}${r.czk.advances ? " − avans " + czk(r.czk.advances) : ""}${r.deductions ? " − odbici " + eur(r.deductions) : ""}${r.czk.deductions ? " − odbici " + czk(r.czk.deductions) : ""} = <b>za isplatu ${[r.net !== 0 || r.czk.net === 0 ? eur(r.net) : "", r.czk.net !== 0 ? czk(r.czk.net) : ""].filter(Boolean).join(" i ")}</b></div>
      <div class="muted" style="margin-top:24px">Potpis radnika: ______________________ &nbsp;&nbsp; Potpis poslodavca: ______________________</div>`);
  };

  /* PDF specifikacija za hotel */
  const printInvoiceSpec = (key, v) => {
    const st = api.settings;
    const ob = data.objects.find((o) => o.id === key);
    const ls = periodLogs.filter((l) => l.objectId === key).sort((a, b) => a.date.localeCompare(b.date));
    const rowsHtml = ls.map((l) => {
      const wk = data.workers.find((x) => x.id === l.workerId);
      return `<tr><td>${fmtDate(l.date)}</td><td>${wk?.name || ""}</td><td>${logSpan(l)}</td><td class="right">${fmtH(l.hours)}</td></tr>`;
    }).join("");
    printDoc(`Specifikacija ${ob?.name}`, `
      ${firmHead(st)}
      <h1>Specifikacija sati — ${ob?.name}</h1><h2>${periodLabel}</h2>
      <table><tr><th>Datum</th><th>Radnik</th><th>Vrijeme</th><th class="right">Sati</th></tr>${rowsHtml}
      <tr class="tot"><td colspan="3">UKUPNO</td><td class="right">${fmtH(v.hours)}</td></tr></table>
      <table><tr><th></th><th class="right">Iznos</th></tr>
      <tr><td>${fmtH(v.hours)} × ${money(ob?.billRate || 0, ob?.billCur)}/h</td><td class="right">${money((ob?.billCur === "CZK" ? v.revenueKc : v.revenue), ob?.billCur)}</td></tr>
      <tr class="tot"><td>ZA NAPLATU</td><td class="right">${money((ob?.billCur === "CZK" ? v.revenueKc : v.revenue), ob?.billCur)}</td></tr></table>
      ${st.iban ? `<div class="muted">Uplata na IBAN: <b>${st.iban}</b></div>` : ""}`);
  };

  /* ---------- Koverte za isplatu (jedna po radniku, za ispis) ---------- */
  const printEnvelopes = () => {
    const pages = rows.map((r) => {
      const objs = [...new Set(r.logs.map((l) => objName(l.objectId)).filter(Boolean))];
      const objLine = objs.length ? objs.join(", ") : (objName(r.w.objectId) || "");
      const amounts = [
        r.net !== 0 || r.czk.net === 0 ? eur(r.net) : "",
        r.czk.net !== 0 ? czk(r.czk.net) : "",
      ].filter(Boolean);
      const gc = wCur(r.w);
      const breakdown = [
        gc === "CZK" ? (r.grossKc ? "zarada " + czk(r.grossKc) : "") : (r.gross ? "zarada " + eur(r.gross) : ""),
        r.bonuses ? "bonus +" + eur(r.bonuses) : "", r.czk.bonuses ? "bonus +" + czk(r.czk.bonuses) : "",
        r.advances ? "avans −" + eur(r.advances) : "", r.czk.advances ? "avans −" + czk(r.czk.advances) : "",
        r.deductions ? "odbici −" + eur(r.deductions) : "", r.czk.deductions ? "odbici −" + czk(r.czk.deductions) : "",
      ].filter(Boolean).join(" · ");
      return `<div class="env">
        <div class="env-top">${periodLabel}${objLine ? "<br>" + objLine : ""}</div>
        <div class="env-mid">
          <div class="env-name">${r.w.name}</div>
          <div class="env-hours">${fmtH(r.hours)}</div>
        </div>
        ${breakdown && breakdown !== ("zarada " + amounts[0]) ? `<div class="env-break">${breakdown}</div>` : ""}
        <div class="env-total">${amounts.join("<br>")}</div>
        <div class="env-space"></div>
      </div>`;
    }).join("");
    printDoc("Koverte " + periodLabel + (objFilterName ? " — " + objFilterName : ""), `<style>
      @page{size:220mm 110mm;margin:0}
      body{padding:0!important;margin:0!important;max-width:none!important}
      .env{width:220mm;height:110mm;box-sizing:border-box;padding:8mm 18mm 6mm;page-break-after:always;
        text-align:center;display:flex;flex-direction:column}
      .env-top{font-size:12pt;color:#444;line-height:1.35}
      .env-mid{margin:3mm 0 2mm}
      .env-name{font-size:21pt;font-weight:bold;letter-spacing:.5px;margin-bottom:1mm;text-transform:uppercase}
      .env-hours{font-size:12.5pt}
      .env-break{font-size:10pt;color:#444;margin-bottom:1.5mm}
      .env-total{border-top:2px solid #000;width:70%;margin:0 auto;padding-top:2.5mm;font-size:19pt;font-weight:bold;line-height:1.25}
      .env-space{flex:1;min-height:24mm}
    </style>${pages}`);
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const obrRows = rows.map((r) => ({
      "Radnik": r.w.name, "Objekt": objName(r.w.objectId) || "", "Sati": r.hours,
      "Satnica": r.rateSet.length === 1 ? r.rateSet[0] : r.rateSet.length === 0 ? rateNow(data, r.w) : "razne",
      "Valuta satnice": wCur(r.w) === "CZK" ? "Kč" : "€",
      "Zarada (€)": r.gross, "Zarada (Kč)": r.grossKc, "Bonus (€)": r.bonuses, "Avans (€)": r.advances, "Odbici (€)": r.deductions,
      "ZA ISPLATU (€)": r.net, "ZA ISPLATU (Kč)": r.czk.net, "Trošak firme (€)": r.firmCosts, "Trošak firme (Kč)": r.czk.firmCosts,
      "Isplaćeno": view === "month" && paidFor(data, r.w.id, month) ? "DA" : "",
    }));
    obrRows.push({ "Radnik": "UKUPNO", "Sati": totals.hours, "Zarada (€)": totals.gross, "Bonus (€)": totals.bonus,
      "ZA ISPLATU (€)": totals.net, "ZA ISPLATU (Kč)": totals.netKc, "Trošak firme (€)": totals.firm, "Trošak firme (Kč)": totals.firmKc });
    const ws1 = XLSX.utils.json_to_sheet(obrRows);
    ws1["!cols"] = [{wch:22},{wch:18},{wch:8},{wch:11},{wch:12},{wch:10},{wch:11},{wch:11},{wch:15},{wch:15},{wch:10}];
    XLSX.utils.book_append_sheet(wb, ws1, "Obračun");

    const satiRows = periodLogs.sort((a, b) => a.date.localeCompare(b.date)).map((l) => {
      const wk = data.workers.find((x) => x.id === l.workerId);
      return { "Datum": fmtDate(l.date), "Radnik": wk?.name || "", "Objekt": objName(l.objectId) || "",
        "Od": l.from || "", "Do": l.to || "", "Sati": l.hours,
        "Satnica": rateFor(data, wk, l.date), "Iznos": round2(l.hours * rateFor(data, wk, l.date)),
        "Valuta": wCur(wk) === "CZK" ? "Kč" : "€" };
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(satiRows.length ? satiRows : [{ "Datum": "" }]), "Sati");

    const payRows = data.payments.filter((pp) => (view === "month" ? inMonth : inYear)(pp.date)).sort((a, b) => a.date.localeCompare(b.date)).map((pp) => ({
      "Datum": fmtDate(pp.date),
      "Za": pp.workerId ? (data.workers.find((x) => x.id === pp.workerId)?.name || "") : "Objekt: " + (objName(pp.objectId) || ""),
      "Vrsta": TYPE_LABEL[pp.type] || pp.type,
      "Iznos": pp.amount,
      "Valuta": pp.currency === "CZK" ? "Kč" : "€",
      "Napomena": pp.note || "",
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(payRows.length ? payRows : [{ "Datum": "" }]), "Isplate i troškovi");

    const objRows = [...byObject.entries()].map(([k, v]) => ({
      "Objekt": k === "__none__" ? "Bez objekta" : objName(k), "Sati": v.hours,
      "Zarada radnika (€)": v.gross, "Zarada radnika (Kč)": v.grossKc || 0,
      "Troškovi objekta (€)": v.costs || 0, "Troškovi objekta (Kč)": v.costsKc || 0,
      ...(admin ? {
        "Naplata (€)": v.revenue, "Naplata (Kč)": v.revenueKc || 0,
        "Dobit (€)": round2(v.revenue - v.gross - (v.costs || 0)),
        "Dobit (Kč)": round2((v.revenueKc || 0) - (v.grossKc || 0) - (v.costsKc || 0)),
      } : {}),
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(objRows.length ? objRows : [{ "Objekt": "" }]), "Po objektima");
    XLSX.writeFile(wb, `Obracun_${view === "month" ? MONTHS[m - 1] + "_" + y : "Godina_" + y}${objFilterName ? "_" + objFilterName.replace(/[^\w]+/g, "") : ""}.xlsx`);
  };

  return (
    <>
      <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
        {[["month", api.t("monthBtn")], ["year", api.t("yearBtn")]].map(([id, label]) => (
          <button key={id} onClick={() => setView(id)} style={{ flex: 1, padding: "9px 8px", borderRadius: 10, fontWeight: 700, fontSize: 13.5, cursor: "pointer",
            background: view === id ? S.ink : "#fff", color: view === id ? "#fff" : S.sub, border: `1px solid ${view === id ? S.ink : S.line}` }}>{label}</button>
        ))}
      </div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
        <Btn small kind="ghost" onClick={view === "month" ? prev : () => setMonth(`${y - 1}-${String(m).padStart(2, "0")}`)}>←</Btn>
        <div style={{ fontWeight: 800, fontSize: 16.5 }}>{periodLabel}</div>
        <Btn small kind="ghost" onClick={view === "month" ? next : () => setMonth(`${y + 1}-${String(m).padStart(2, "0")}`)}>→</Btn>
      </div>

      <Field label={api.t("showFor")}>
        <select value={objFilter} onChange={(e) => setObjFilter(e.target.value)}>
          <option value="">{api.t("allObjects")}</option>
          {["HR", "CZ"].map((c) => {
            const objs = sortedObjects(data.objects.filter((o) => (o.country || "HR") === c));
            return objs.length ? (
              <optgroup key={c} label={`${FLAG[c]} ${COUNTRY_NAME[c]}`}>
                {objs.map((o) => <option key={o.id} value={o.id}>{o.name}</option>)}
              </optgroup>
            ) : null;
          })}
        </select>
      </Field>

      {rows.length === 0 ? <Empty text={objFilter ? `Za ${objFilterName} u ovom razdoblju nema upisanih sati.` : "Za ovo razdoblje nema upisanih sati ni isplata."} /> : (
        <>
          <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
            <Btn kind="excel" onClick={exportExcel} style={{ flex: 1 }}>📊 {api.t("excelBtn")}{objFilterName ? " — " + objFilterName : ""}</Btn>
            <Btn kind="ghost" onClick={printEnvelopes} style={{ flex: 1, fontWeight: 700 }}>🖨 {api.t("envelopesBtn")}{objFilterName ? " — " + objFilterName : ""}</Btn>
          </div>

          <Card style={{ background: S.ink, color: "#fff", border: "none" }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, fontSize: 13.5 }}>
              <div><div style={{ color: "#9DB3A8" }}>{api.t("totalHours")}</div><div className="num" style={{ fontSize: 18, fontWeight: 800 }}>{fmtH(totals.hours)}</div></div>
              <div><div style={{ color: "#9DB3A8" }}>{api.t("earningsBonus")}</div>
                <div className="num" style={{ fontSize: 18, fontWeight: 800 }}>{eur(round2(totals.gross + totals.bonus))}</div>
                {round2(totals.grossKc + totals.bonusKc) !== 0 && <div className="num" style={{ fontSize: 13.5, fontWeight: 700 }}>{czk(round2(totals.grossKc + totals.bonusKc))}</div>}
              </div>
              <div><div style={{ color: "#9DB3A8" }}>{api.t("toPay")}</div>
                <div className="num" style={{ fontSize: 18, fontWeight: 800, color: "#7FD6B4" }}>{eur(totals.net)}</div>
                {totals.netKc !== 0 && <div className="num" style={{ fontSize: 13.5, fontWeight: 700, color: "#7FD6B4" }}>{totals.netKc > 0 ? "" : "−"}{czk(Math.abs(totals.netKc))} u Kč</div>}
              </div>
              <div><div style={{ color: "#9DB3A8" }}>{api.t("firmCosts")}</div>
                <div className="num" style={{ fontSize: 18, fontWeight: 800 }}>{eur(round2(totals.firm + objCostsTotal))}</div>
                {round2(totals.firmKc + objCostsTotalKc) !== 0 && <div className="num" style={{ fontSize: 13.5, fontWeight: 700 }}>{czk(round2(totals.firmKc + objCostsTotalKc))}</div>}
              </div>
            </div>
          </Card>

          {/* grafovi */}
          <Card>
            <div style={{ fontWeight: 700, marginBottom: 4 }}>📈 Sati — zadnjih 6 mjeseci</div>
            <MiniBars items={last6.map((x) => ({ label: x.label, value: x.hours }))} fmt={(v) => Math.round(v)} color={S.blue} />
            {admin && (
              <>
                <div style={{ fontWeight: 700, margin: "12px 0 4px", color: S.amber }}>👑 Dobit u € — zadnjih 6 mjeseci</div>
                <MiniBars items={last6.map((x) => ({ label: x.label, value: x.profit }))} fmt={(v) => Math.round(v) + "€"} color={S.green} />
                {last6.some((x) => x.profitKc !== 0) && (
                  <>
                    <div style={{ fontWeight: 700, margin: "12px 0 4px", color: S.amber }}>👑 Dobit u Kč — zadnjih 6 mjeseci</div>
                    <MiniBars items={last6.map((x) => ({ label: x.label, value: x.profitKc }))} fmt={(v) => Math.round(v) + " Kč"} color={S.green} />
                  </>
                )}
              </>
            )}
          </Card>

          {admin && (
            <Card style={{ background: S.amberSoft, borderColor: "#EBD9B4" }}>
              <div style={{ fontWeight: 700, color: S.amber, marginBottom: 8 }}>👑 Admin — dobit ({periodLabel})</div>
              <div className="num" style={{ fontSize: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>Naplata od objekata</span><b>{eur(totalRevenue)}{totalRevenueKc !== 0 ? " + " + czk(totalRevenueKc) : ""}</b></div>
                <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>Plaće + bonusi + svi troškovi</span><b>−{eur(totalCost)}{totalCostKc !== 0 ? " − " + czk(totalCostKc) : ""}</b></div>
                <div style={{ borderTop: `1px dashed #EBD9B4`, paddingTop: 6 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 800, fontSize: 15.5 }}>
                    <span>DOBIT (€)</span><span style={{ color: totalProfit >= 0 ? S.green : S.red }}>{eur(totalProfit)}</span>
                  </div>
                  {(totalRevenueKc !== 0 || totalCostKc !== 0) && (
                    <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 800, fontSize: 15.5 }}>
                      <span>DOBIT (Kč)</span><span style={{ color: totalProfitKc >= 0 ? S.green : S.red }}>{czk(totalProfitKc)}</span>
                    </div>
                  )}
                </div>
              </div>
            </Card>
          )}

          {/* po objektima + naplata od hotela */}
          {byObject.size > 0 && (
            <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA" }}>
              <div style={{ fontWeight: 700, color: S.blue, marginBottom: 8 }}>🏨 Po objektima</div>
              {["HR", "CZ", "__none__"].map((cGroup) => {
                const entries = [...byObject.entries()].filter(([key]) => {
                  if (cGroup === "__none__") return key === "__none__";
                  const ob = data.objects.find((o) => o.id === key);
                  return ob && (ob.country || "HR") === cGroup;
                });
                if (!entries.length) return null;
                const sub = entries.reduce((t, [, v]) => ({
                  hours: round2(t.hours + v.hours), gross: round2(t.gross + v.gross), grossKc: round2(t.grossKc + (v.grossKc || 0)),
                  revenue: round2(t.revenue + v.revenue), revenueKc: round2(t.revenueKc + (v.revenueKc || 0)),
                  costs: round2(t.costs + (v.costs || 0)), costsKc: round2(t.costsKc + (v.costsKc || 0)),
                }), { hours: 0, gross: 0, grossKc: 0, revenue: 0, revenueKc: 0, costs: 0, costsKc: 0 });
                const subProfitE = round2(sub.revenue - sub.gross - sub.costs);
                const subProfitK = round2(sub.revenueKc - sub.grossKc - sub.costsKc);
                return (
                  <div key={cGroup}>
                    {cGroup !== "__none__" && (
                      <div className="num" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0 3px", fontWeight: 800, fontSize: 13.5, color: S.blue }}>
                        <span>{FLAG[cGroup]} {COUNTRY_NAME[cGroup]}</span>
                        <span>{fmtH(sub.hours)}{admin ? <> · dobit {[sub.revenue || sub.gross || sub.costs ? eur(subProfitE) : "", sub.revenueKc || sub.grossKc || sub.costsKc ? czk(subProfitK) : ""].filter(Boolean).join(" · ") || eur(0)}</> : null}</span>
                      </div>
                    )}
                    {entries.map(([key, v]) => {
                const ups = view === "month" ? (data.invoicePayments || []).filter((p) => p.objectId === key && p.month === month) : [];
                const paid = round2(ups.reduce((s, p) => s + p.amount, 0));
                const bcKey = data.objects.find((o) => o.id === key)?.billCur || "EUR";
                const due = round2((bcKey === "CZK" ? v.revenueKc : v.revenue) - paid);
                return (
                  <div key={key} style={{ padding: "7px 0", borderBottom: `1px solid #CBDCEA` }}>
                    <div className="num" style={{ display: "flex", justifyContent: "space-between", fontSize: 14 }}>
                      <span style={{ fontWeight: 600 }}>{key === "__none__" ? "Bez objekta" : objName(key)}</span>
                      <span>{fmtH(v.hours)} · <b>{[v.gross ? eur(v.gross) : "", v.grossKc ? czk(v.grossKc) : ""].filter(Boolean).join(" + ") || eur(0)}</b>{(v.costs > 0 || v.costsKc > 0) ? <span style={{ color: S.red }}> · troškovi {[v.costs ? eur(v.costs) : "", v.costsKc ? czk(v.costsKc) : ""].filter(Boolean).join(" + ")}</span> : null}</span>
                    </div>
                    {admin && key !== "__none__" && (
                      <>
                        <div className="num" style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, color: S.amber, fontWeight: 700 }}>
                          <span>naplata {[v.revenue ? eur(v.revenue) : "", v.revenueKc ? czk(v.revenueKc) : ""].filter(Boolean).join(" + ") || eur(0)}</span>
                          <span>dobit {[
                            (v.revenue || v.gross || v.costs) ? eur(round2(v.revenue - v.gross - (v.costs || 0))) : "",
                            (v.revenueKc || v.grossKc || v.costsKc) ? czk(round2((v.revenueKc || 0) - (v.grossKc || 0) - (v.costsKc || 0))) : "",
                          ].filter(Boolean).join(" · ") || eur(0)}</span>
                        </div>
                        {view === "month" && (
                          <div style={{ marginTop: 6 }}>
                            <div className="num" style={{ fontSize: 12.5, display: "flex", justifyContent: "space-between" }}>
                              <span>uplaćeno: <b style={{ color: S.green }}>{money(paid, bcKey)}</b></span>
                              <span>ostaje: <b style={{ color: due > 0 ? S.red : S.green }}>{money(Math.max(due, 0), bcKey)}</b></span>
                            </div>
                            {ups.map((p) => (
                              <div key={p.id} className="num" style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: S.sub, padding: "2px 0" }}>
                                <span>{fmtDate(p.date)} · uplata {money(p.amount, bcKey)}</span>
                                <button onClick={() => api.delInvoicePayment(p, objName(key))} style={{ background: "none", border: "none", color: S.red, cursor: "pointer" }}>✕</button>
                              </div>
                            ))}
                            <div style={{ display: "flex", gap: 6, marginTop: 4 }}>
                              <input inputMode="decimal" placeholder={bcKey === "CZK" ? "iznos uplate Kč" : "iznos uplate €"} value={uplata[key] || ""}
                                onChange={(e) => setUplata({ ...uplata, [key]: e.target.value })} style={{ padding: "7px 9px", fontSize: 13 }} />
                              <Btn small kind="ghost" onClick={() => {
                                const a = parseNum(uplata[key]);
                                if (a) { api.addInvoicePayment(data.objects.find((o) => o.id === key), month, round2(a)); setUplata({ ...uplata, [key]: "" }); }
                              }}>+ Uplata</Btn>
                              <Btn small kind="ghost" onClick={() => printInvoiceSpec(key, v)}>🧾 PDF</Btn>
                            </div>
                          </div>
                        )}
                      </>
                    )}
                  </div>
                );
                    })}
                  </div>
                );
              })}
            </Card>
          )}

          {view === "month" && rows.length > 0 && (
            <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 4px", marginBottom: 4, flexWrap: "wrap" }}>
              {unpaidRows.length > 0 && (
                <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, cursor: "pointer" }}>
                  <input type="checkbox" checked={allSelected} onChange={toggleSelectAll} style={{ width: 18, height: 18 }} />
                  Odaberi sve neisplaćene ({unpaidRows.length})
                </label>
              )}
              <div style={{ display: "flex", gap: 8, marginLeft: "auto" }}>
                {selected.size > 0 && (
                  <Btn small onClick={() => payMany(rows.filter((r) => selected.has(r.w.id)))} disabled={bulkBusy}>
                    ✓ Isplati odabrane ({selected.size})
                  </Btn>
                )}
                {unpaidRows.length > 0 && (
                  <Btn small kind={selected.size > 0 ? "ghost" : "primary"} onClick={() => payMany(unpaidRows)} disabled={bulkBusy}>
                    {bulkBusy ? "Isplaćujem…" : `✓✓ Isplati sve (${unpaidRows.length})`}
                  </Btn>
                )}
              </div>
            </div>
          )}

          {rows.map((r) => {
            const paid = view === "month" ? paidFor(data, r.w.id, month) : null;
            return (
              <Card key={r.w.id} style={paid ? { borderColor: "#C5DED2", background: "#FBFDF9" } : undefined}>
                <div onClick={() => setOpen(open === r.w.id ? null : r.w.id)} style={{ cursor: "pointer" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div style={{ fontWeight: 700, fontSize: 15.5, display: "flex", alignItems: "center", gap: 8 }}>
                      {view === "month" && !paid && (
                        <input type="checkbox" checked={selected.has(r.w.id)} onChange={(e) => toggleSelect(r.w.id, e)} onClick={(e) => e.stopPropagation()} style={{ width: 18, height: 18 }} />
                      )}
                      {r.w.name} {paid && <Tag color={S.green} bg={S.greenSoft}>✓ Isplaćeno</Tag>}
                    </div>
                    <div className="num" style={{ textAlign: "right" }}>
                      <div style={{ fontWeight: 800, color: r.net >= 0 ? S.green : S.red, fontSize: 16 }}>{eur(r.net)}</div>
                      {r.czk.net !== 0 && <div style={{ fontWeight: 700, color: r.czk.net >= 0 ? S.green : S.red, fontSize: 12.5 }}>{r.czk.net > 0 ? "+" : "−"}{czk(Math.abs(r.czk.net))}</div>}
                    </div>
                  </div>
                  <div className="num" style={{ fontSize: 13, color: S.sub, marginTop: 3 }}>
                    {fmtH(r.hours)}{r.rateSet.length === 1 ? <> × {money(r.rateSet[0], wCur(r.w))}</> : r.rateSet.length > 1 ? <> (više satnica)</> : null} = {wCur(r.w) === "CZK" ? czk(r.grossKc) : eur(r.gross)}
                    {r.bonuses > 0 && <> · bonus +{eur(r.bonuses)}</>}
                    {r.advances > 0 && <> · avans −{eur(r.advances)}</>}
                    {r.deductions > 0 && <> · odbici −{eur(r.deductions)}</>}
                    {r.czk.advances > 0 && <> · avans −{czk(r.czk.advances)}</>}
                    {r.czk.bonuses > 0 && <> · bonus +{czk(r.czk.bonuses)}</>}
                    {r.czk.deductions > 0 && <> · odbici −{czk(r.czk.deductions)}</>}
                  </div>
                </div>
                {open === r.w.id && (
                  <div style={{ marginTop: 12, borderTop: `1px dashed ${S.line}`, paddingTop: 10 }}>
                    {r.logs.sort((a, b) => a.date.localeCompare(b.date)).map((l) => (
                      <div key={l.id} className="num" style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, padding: "4px 0" }}>
                        <span>{fmtDate(l.date)} · {logSpan(l)}{objName(l.objectId) ? " · " + objName(l.objectId) : ""}</span>
                        <span style={{ fontWeight: 600 }}>{fmtH(l.hours)}</span>
                      </div>
                    ))}
                    {r.pays.map((p) => (
                      <div key={p.id} className="num" style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, padding: "4px 0",
                        color: p.type === "bonus" ? S.green : p.deduct ? S.amber : S.sub }}>
                        <span>{fmtDate(p.date)} · {TYPE_LABEL[p.type] || p.type}{p.note ? " · " + p.note : ""}</span>
                        <span style={{ fontWeight: 600 }}>{p.type === "bonus" ? "+" : p.deduct ? "−" : ""}{money(p.amount, p.currency)}</span>
                      </div>
                    ))}
                    <div style={{ borderTop: `1px dashed ${S.line}`, marginTop: 8, paddingTop: 8, display: "flex", justifyContent: "space-between", fontWeight: 800 }}>
                      <span>Za isplatu</span>
                      <span className="num" style={{ color: r.net >= 0 ? S.green : S.red, textAlign: "right" }}>
                        {eur(r.net)}{r.czk.net !== 0 ? <span style={{ display: "block", fontSize: 13 }}>{r.czk.net > 0 ? "+" : "−"}{czk(Math.abs(r.czk.net))}</span> : null}
                      </span>
                    </div>
                    <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
                      <Btn small kind="ghost" onClick={() => printPayslip(r)}>📄 PDF obračun</Btn>
                      {view === "month" && !paid && (
                        <Btn small onClick={() => api.markPaid(r.w, month, r.net, r.czk.net)}>✓ Označi isplaćeno</Btn>
                      )}
                      {view === "month" && paid && (
                        <>
                          <span className="num" style={{ fontSize: 12.5, color: S.green, alignSelf: "center", fontWeight: 700 }}>
                            Isplaćeno {fmtDate(paid.paidAt)} · {[paid.amount ? eur(paid.amount) : "", paid.amountKc ? czk(paid.amountKc) : ""].filter(Boolean).join(" + ") || eur(0)}
                          </span>
                          {admin && <Btn small kind="danger" onClick={() => api.unmarkPaid(paid, r.w.name)}>Otključaj</Btn>}
                        </>
                      )}
                    </div>
                    {view === "month" && paid && (
                      <div style={{ fontSize: 11.5, color: S.sub, marginTop: 6 }}>Mjesec je zaključan — sati i isplate se više ne mogu mijenjati{admin ? "" : " (otključava admin)"}.</div>
                    )}
                  </div>
                )}
              </Card>
            );
          })}
        </>
      )}
    </>
  );
}
