/* ================================================================== */
/*  FINANCIJE — bankovni izvodi, kategorije i procjena poreza (CZ s.r.o.) */
/*  Čista logika bez Reacta: parsiranje izvoda, pravila, zbrojevi.       */
/* ================================================================== */

// Stope za Češku (2026.). Procjena — konačan iznos uvijek potvrdi s knjigovođom.
export const CZ_CIT = 0.21;        // daň z příjmů právnických osob
export const CZ_DIVIDEND = 0.15;   // srážková daň z dividend (fizička osoba, rezident)
// odvodi iz plaće (2026.): sociální 24,8 % firma + 7,1 % zaposlenik, zdravotní 9 % firma + 4,5 % zaposlenik
export const CZ_SOC_FIRM = 0.248, CZ_SOC_EMP = 0.071, CZ_ZDR_FIRM = 0.09, CZ_ZDR_EMP = 0.045;

// [ključ, naziv, vrsta, zadana stopa DPH u %]
// vrsta: in = prihod, cost = porezno priznat trošak, tax = plaćeni porezi (nisu trošak),
//        owner = isplata vlasniku (nije trošak), neutral = ne ulazi u dobit (prijenosi, zajmovi)
export const FIN_CATS = [
  ["prihod", "💶 Prihod od klijenata", "in", 21],
  ["prihod_ostalo", "➕ Ostali prihodi", "in", 0],
  ["place", "👷 Plaće radnicima", "cost", 0],
  ["odvodi", "🏥 Socijalno i zdravstveno (ČSSZ, VZP…)", "cost", 0],
  ["porez_place", "🧾 Porez iz plaća (záloha na daň)", "cost", 0],
  ["smjestaj", "🛏 Smještaj radnika", "cost", 12],
  ["gorivo", "⛽ Gorivo i auto", "cost", 21],
  ["najam", "🏠 Najam", "cost", 21],
  ["usluge", "📚 Knjigovođa i usluge", "cost", 21],
  ["oprema", "🧰 Oprema i materijal", "cost", 21],
  ["banka", "🏦 Bankovne naknade", "cost", 0],
  ["ostali_troskovi", "📦 Ostali troškovi", "cost", 21],
  ["dph", "🧾 Plaćeni DPH", "tax", 0],
  ["porez_dobit", "🏛 Porez na dobit (zálohy)", "tax", 0],
  ["vlasnik", "👤 Isplata vlasniku / dividenda", "owner", 0],
  ["prijenos", "🔁 Prijenos između računa", "neutral", 0],
  ["zajam", "💳 Zajam / povrat zajma", "neutral", 0],
];
export const FIN_CAT = Object.fromEntries(FIN_CATS.map(([key, label, kind, vat]) => [key, { key, label, kind, vat }]));
export const catLabel = (k) => FIN_CAT[k]?.label || "❔ Nerazvrstano";
export const catKind = (k) => FIN_CAT[k]?.kind || "";

// Ugrađena pravila (rade i bez ičega što admin upiše). Uspoređuje se bez kvačica i velikih slova.
export const DEFAULT_RULES = [
  ["poplatek", "banka"], ["vedeni uctu", "banka"], ["vedení účtu", "banka"], ["bank fee", "banka"], ["urok", "prihod_ostalo"],
  ["financni urad", "dph"], ["fu pro", "dph"],
  ["cssz", "odvodi"], ["ossz", "odvodi"], ["socialni zabezpeceni", "odvodi"], ["vzp", "odvodi"], ["zdravotni pojistovna", "odvodi"],
  ["ozp", "odvodi"], ["cpzp", "odvodi"], ["vozp", "odvodi"], ["zpmv", "odvodi"], ["zp mv", "odvodi"], ["rbp", "odvodi"],
  ["mzda", "place"], ["mzdy", "place"], ["vyplata", "place"], ["plat", "place"],
  ["benzina", "gorivo"], ["shell", "gorivo"], ["omv", "gorivo"], ["orlen", "gorivo"], ["mol", "gorivo"], ["eurooil", "gorivo"], ["tank ono", "gorivo"],
  ["najem", "najam"], ["ubytovani", "smjestaj"], ["ubytovna", "smjestaj"],
  ["ucetni", "usluge"], ["ucetnictvi", "usluge"],
  ["dividend", "vlasnik"], ["podil na zisku", "vlasnik"],
  ["prevod mezi ucty", "prijenos"], ["vlastni ucet", "prijenos"],
  ["/0710", "dph"], // račun poreznog ureda kod ČNB (ČSSZ/VZP su gore pa imaju prednost)
];

export const fold = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim();
export const txText = (tx) => fold([tx.counterparty, tx.counter_account, tx.vs, tx.message].join(" "));
// riječ mora počinjati s uzorkom; kratki uzorci (do 4 znaka, npr. "plat", "mol") moraju biti cijela riječ
export const ruleMatches = (text, pattern) => {
  const p = fold(pattern);
  if (!p) return false;
  const esc = p.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`${/^[a-z0-9]/.test(p) ? "(?:^|[^a-z0-9])" : ""}${esc}${p.length <= 4 ? "(?![a-z0-9])" : ""}`).test(text);
};

// Vraća ključ kategorije prema pravilima (prvo adminova, onda ugrađena); prihod je zadan za uplate.
export function suggestCategory(tx, userRules = []) {
  const text = txText(tx);
  for (const r of userRules) if (ruleMatches(text, r.pattern)) return r.category;
  for (const [p, cat] of DEFAULT_RULES) {
    if (!ruleMatches(text, p)) continue;
    if (FIN_CAT[cat].kind === "in" && tx.amount < 0) continue;
    if (FIN_CAT[cat].kind !== "in" && cat !== "dph" && tx.amount > 0) continue;
    return cat;
  }
  return tx.amount > 0 ? "prihod" : "";
}

/* ---------- čitanje datoteke ---------- */
export function decodeText(buf) {
  try { return new TextDecoder("utf-8", { fatal: true }).decode(buf).replace(/^﻿/, ""); }
  catch { return new TextDecoder("windows-1250").decode(buf); }
}

export function parseCsv(text) {
  const lines = text.split(/\r?\n/).slice(0, 40);
  const count = (ch) => lines.reduce((s, l) => s + (l.split(ch).length - 1), 0);
  const delim = [";", "\t", ","].sort((a, b) => count(b) - count(a))[0];
  const rows = [];
  let row = [], cell = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"' && text[i + 1] === '"') { cell += '"'; i++; }
      else if (c === '"') q = false;
      else cell += c;
    } else if (c === '"') q = true;
    else if (c === delim) { row.push(cell); cell = ""; }
    else if (c === "\n" || c === "\r") {
      if (c === "\r" && text[i + 1] === "\n") i++;
      row.push(cell); rows.push(row); row = []; cell = "";
    } else cell += c;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  return rows.filter((r) => r.some((x) => String(x).trim()));
}

/* ---------- prepoznavanje stupaca (Fio, KB, ČSOB, Česká spořitelna, Raiffeisen, Air Bank, Moneta, UniCredit…) ---------- */
export const FIELDS = [
  ["date", "Datum"], ["amount", "Iznos"], ["credit", "Uplata (ako je odvojeno)"], ["debit", "Isplata (ako je odvojeno)"],
  ["currency", "Valuta"], ["counterparty", "Partner"], ["account", "Račun partnera"], ["bankCode", "Kód banky"],
  ["vs", "VS"], ["direction", "Smjer"], ["txid", "ID transakcije"],
];
const PATTERNS = {
  date: [/^datum zauctovani/, /^datum$/, /^datum provedeni/, /^datum pohybu/, /^datum transakce/, /^datum splatnosti/, /booking date/, /^date$/, /datum/],
  amount: [/^objem$/, /^castka$/, /^zauctovana castka/, /castka v mene uctu/, /^amount/, /^castka/, /^iznos/],
  credit: [/^kredit/, /^credit/, /^prijem/, /^prichozi/],
  debit: [/^debet/, /^debit/, /^vydaj/, /^odchozi/],
  currency: [/^mena$/, /^mena uctu/, /^mena/, /currency/],
  counterparty: [/nazev protiuctu/, /nazev protistrany/, /nazev uctu protiuctu/, /^protistrana$/, /^partner/, /nazev prijemce/, /counterparty/, /^prijemce/, /^nazev obchodnika/, /^obchodnik/],
  account: [/^protiucet/, /cislo uctu protistrany/, /cislo protiuctu/, /^cislo uctu protiuctu/, /ucet protistrany/, /counter.*account/, /^iban protistrany/],
  bankCode: [/kod banky/],
  vs: [/^vs$/, /variabilni symbol/, /variable symbol/],
  direction: [/^smer/],
  txid: [/^id pohybu/, /^id transakce/, /identifikace transakce/, /transaction id/, /^id operace/],
};
const MSG_RE = /zprava|popis|poznamka|message|note|oznaceni operace|^typ|details|reference|ucel|^ks$|^ss$/;

export function findHeader(rows) {
  for (let i = 0; i < Math.min(rows.length, 40); i++) {
    const h = rows[i].map(fold);
    const has = (f) => h.some((c) => PATTERNS[f].some((re) => re.test(c)));
    if (has("date") && (has("amount") || has("credit"))) return i;
  }
  return -1;
}

export function guessMapping(header) {
  const h = header.map(fold);
  const used = new Set();
  const map = {};
  for (const [f] of FIELDS) {
    for (const re of PATTERNS[f]) {
      const i = h.findIndex((c, j) => !used.has(j) && re.test(c));
      if (i >= 0) { map[f] = i; used.add(i); break; }
    }
  }
  map.message = h.map((c, j) => (!used.has(j) && MSG_RE.test(c) ? j : -1)).filter((j) => j >= 0);
  return map;
}

export function parseAmount(v) {
  if (typeof v === "number") return v;
  let s = String(v || "").replace(/[\s  ']/g, "").replace(/[−–]/g, "-").replace(/[^\d.,+-]/g, "");
  if (!s) return NaN;
  let neg = false;
  if (s.endsWith("-")) { neg = true; s = s.slice(0, -1); }
  const lc = s.lastIndexOf(","), ld = s.lastIndexOf(".");
  if (lc >= 0 && ld >= 0) s = lc > ld ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  else if (lc >= 0) s = s.replace(",", ".");
  const n = parseFloat(s);
  return neg ? -n : n;
}

export function parseDate(v) {
  if (v instanceof Date && !isNaN(v)) {
    const d = new Date(v.getTime() - v.getTimezoneOffset() * 60000);
    return d.toISOString().slice(0, 10);
  }
  const s = String(v || "").trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  m = s.match(/^(\d{1,2})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{2,4})/);
  if (m) {
    const y = m[3].length === 2 ? "20" + m[3] : m[3];
    return `${y}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  }
  return "";
}

// kratki hash (cyrb53) — za prepoznavanje već uvezenih transakcija
function cyrb53(str, seed = 0) {
  let h1 = 0xdeadbeef ^ seed, h2 = 0x41c6ce57 ^ seed;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761); h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}

// Pretvara retke izvoda u transakcije. Vraća { txs, skipped }.
export function rowsToTxs(rows, headerIdx, map, defaultCurrency = "CZK") {
  const cell = (r, i) => (i == null || i < 0 ? "" : String(r[i] ?? "").trim());
  const txs = [];
  let skipped = 0;
  const seen = {};
  for (const r of rows.slice(headerIdx + 1)) {
    const date = parseDate(r[map.date]);
    let amount = map.amount != null && map.amount >= 0 ? parseAmount(r[map.amount]) : NaN;
    if (isNaN(amount) && (map.credit >= 0 || map.debit >= 0)) {
      const cr = parseAmount(r[map.credit]), db = parseAmount(r[map.debit]);
      amount = (isNaN(cr) ? 0 : Math.abs(cr)) - (isNaN(db) ? 0 : Math.abs(db));
    }
    if (!date || isNaN(amount) || amount === 0) { skipped++; continue; }
    if (amount > 0 && /odchoz|outgoing|debit|vydaj/.test(fold(cell(r, map.direction)))) amount = -amount;
    const acc = cell(r, map.account);
    const code = cell(r, map.bankCode);
    const tx = {
      tx_date: date,
      amount: Math.round(amount * 100) / 100,
      currency: (cell(r, map.currency).toUpperCase().match(/CZK|EUR|USD/) || [defaultCurrency])[0],
      counterparty: cell(r, map.counterparty),
      counter_account: acc && code && !acc.includes("/") ? `${acc}/${code}` : acc,
      vs: cell(r, map.vs).replace(/^0+/, ""),
      message: [...new Set((map.message || []).map((i) => cell(r, i)).filter(Boolean))].join(" · ").slice(0, 500),
    };
    const id = cell(r, map.txid);
    const base = id ? `id|${id}` : `${tx.tx_date}|${tx.amount}|${tx.currency}|${tx.counter_account}|${tx.vs}|${tx.counterparty}|${tx.message}`;
    seen[base] = (seen[base] || 0) + 1;
    tx.hash = `${tx.tx_date}|${tx.amount}|${cyrb53(base)}|${seen[base]}`;
    txs.push(tx);
  }
  return { txs, skipped };
}

/* ---------- zbrojevi i procjena poreza ---------- */
export const vatOf = (tx) => (tx.vat_rate == null ? (FIN_CAT[tx.category]?.vat ?? 0) : Number(tx.vat_rate));
export const toCzkAmt = (tx, rate) => (tx.currency === "EUR" ? tx.amount * rate : tx.amount);
export const netOf = (tx, rate) => toCzkAmt(tx, rate) / (1 + vatOf(tx) / 100);

// Zbroj po mjesecima za jednu godinu. Uplate bez kategorije računaju se kao prihod (sigurnija procjena),
// isplate bez kategorije se ne odbijaju dok ih ne razvrstaš.
export function finYear(txs, year, rate) {
  const months = Array.from({ length: 12 }, (_, i) => ({ month: `${year}-${String(i + 1).padStart(2, "0")}`, income: 0, cost: 0, taxPaid: 0, vatPaid: 0, owner: 0, neutral: 0, uncat: 0, count: 0 }));
  const byCat = {};
  for (const tx of txs) {
    if (!tx.tx_date?.startsWith(String(year))) continue;
    const m = months[Number(tx.tx_date.slice(5, 7)) - 1];
    m.count++;
    const cat = tx.category || (tx.amount > 0 ? "prihod" : "");
    const kind = catKind(cat);
    const gross = toCzkAmt(tx, rate), net = netOf({ ...tx, category: cat }, rate);
    (byCat[cat || ""] ||= { gross: 0, net: 0, count: 0 });
    byCat[cat || ""].gross += gross; byCat[cat || ""].net += net; byCat[cat || ""].count++;
    if (!tx.category) m.uncat++;
    if (kind === "in") m.income += net;
    else if (kind === "cost") m.cost -= net;
    else if (cat === "dph") m.vatPaid -= gross;
    else if (cat === "porez_dobit") m.taxPaid -= gross;
    else if (kind === "owner") m.owner -= gross;
    else if (kind === "neutral") m.neutral += gross;
  }
  return { months, byCat };
}

// Procjena poreza na dobit i dividende. base = prihodi − troškovi (bez DPH), zaokruženo na tisuće naniže.
export function taxEstimate({ income, cost, monthsCovered }) {
  const profit = income - cost;
  const base = Math.max(0, Math.floor(profit / 1000) * 1000);
  const cit = Math.round(base * CZ_CIT);
  const factor = monthsCovered > 0 && monthsCovered < 12 ? 12 / monthsCovered : 1;
  const yProfit = profit * factor;
  const yBase = Math.max(0, Math.floor(yProfit / 1000) * 1000);
  const yCit = Math.round(yBase * CZ_CIT);
  const afterTax = Math.max(0, yProfit - yCit);
  return { profit, base, cit, factor, yProfit, yBase, yCit, afterTax, dividendTax: Math.round(afterTax * CZ_DIVIDEND) };
}

/* ---------- godišnji obračun obveza (odvodi iz plaća, rokovi, zálohy) ---------- */
// vrsta odvoda s izvoda: soc = ČSSZ, zdr = zdravstveno osiguranje, dan = porez iz plaća, mix = "odvodi" bez prepoznatog primatelja
export function levyKind(tx) {
  if (tx.category === "porez_place") return "dan";
  if (tx.category !== "odvodi") return "";
  const t = txText(tx);
  if (/cssz|ossz|socialn|zabezpeceni/.test(t)) return "soc";
  if (/vzp|ozp|cpzp|vozp|zpmv|zp mv|rbp|zdravotn|pojistovn/.test(t)) return "zdr";
  return "mix";
}

// podsjetnik iz Mzdy PDF-a ("Platiti … 12 345 Kč – sociální pojištění 05/2026") → { kind, month, amount }
export function parseLevyReminder(text) {
  const m = String(text || "").replace(/\s+/g, " ").match(/^Platiti .*?(?<!\d)(\d{1,3}(?: \d{3})*(?:,\d+)?) Kč – (.+) (\d{2})\/(\d{4})$/);
  if (!m) return null;
  const label = fold(m[2]);
  const kind = /zdrav/.test(label) ? "zdr" : /soci/.test(label) ? "soc" : /dan/.test(label) ? "dan" : "";
  if (!kind) return null;
  return { kind, month: `${m[4]}-${m[3]}`, amount: Number(m[1].replace(/ /g, "").replace(",", ".")) };
}

// Odvodi iz plaća po mjesecu obračuna plaće. Izvor po mjesecu: příkazy iz Mzdy PDF-a (podsjetnici),
// a ako ih za taj mjesec nema, uplate s izvoda (plaćaju se do 20. idućeg mjeseca → mjesec ranije).
export function payrollLevies({ txs = [], reminders = [], year, rate = 25 }) {
  const KINDS = ["soc", "zdr", "dan", "mix"];
  const empty = () => Object.fromEntries(KINDS.map((k) => [k, Array(12).fill(0)]));
  const pdf = empty(), bank = empty();
  for (const r of reminders) {
    const p = parseLevyReminder(r.text);
    if (p && p.month.startsWith(`${year}-`)) pdf[p.kind][Number(p.month.slice(5, 7)) - 1] += p.amount;
  }
  for (const tx of txs) {
    const k = levyKind(tx);
    if (!k || !tx.tx_date?.startsWith(String(year)) || tx.amount > 0) continue;
    const mi = Number(tx.tx_date.slice(5, 7)) - 2; // uplata u veljači = plaće za siječanj
    if (mi >= 0) bank[k][mi] -= toCzkAmt(tx, rate);
  }
  const months = Array.from({ length: 12 }, (_, i) => {
    const usePdf = KINDS.some((k) => pdf[k][i]);
    const src = usePdf ? pdf : bank;
    const row = Object.fromEntries(KINDS.map((k) => [k, src[k][i]]));
    row.source = usePdf ? "pdf" : KINDS.some((k) => bank[k][i]) ? "banka" : "";
    return row;
  });
  const covered = months.filter((m) => m.source).length;
  const total = Object.fromEntries(KINDS.map((k) => [k, months.reduce((s, m) => s + m[k], 0)]));
  const factor = covered > 0 && covered < 12 ? 12 / covered : 1;
  const year12 = Object.fromEntries(KINDS.map((k) => [k, total[k] * factor]));
  // podjela na dio firme i dio zaposlenika (po zakonskim stopama; "mix" se dijeli kao prosjek soc. i zdr.)
  const socFirmShare = CZ_SOC_FIRM / (CZ_SOC_FIRM + CZ_SOC_EMP), zdrFirmShare = CZ_ZDR_FIRM / (CZ_ZDR_FIRM + CZ_ZDR_EMP);
  const mixFirmShare = (CZ_SOC_FIRM + CZ_ZDR_FIRM) / (CZ_SOC_FIRM + CZ_SOC_EMP + CZ_ZDR_FIRM + CZ_ZDR_EMP);
  const split = (t) => {
    const firm = t.soc * socFirmShare + t.zdr * zdrFirmShare + t.mix * mixFirmShare;
    const all = t.soc + t.zdr + t.mix + t.dan;
    return { firm, employees: all - firm, all };
  };
  return { months, covered, factor, total, year12, split: split(total), split12: split(year12) };
}

// Zálohe na porez na dobit prema zadnjoj poznatoj poreznoj obvezi (§ 38a ZDP)
export function citAdvances(lastTax) {
  if (lastTax <= 30000) return { count: 0, amount: 0, dates: [] };
  if (lastTax <= 150000) return { count: 2, amount: Math.ceil(lastTax * 0.4), dates: ["15. 6.", "15. 12."] };
  return { count: 4, amount: Math.ceil(lastTax * 0.25), dates: ["15. 3.", "15. 6.", "15. 9.", "15. 12."] };
}
