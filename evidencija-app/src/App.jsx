import { useState, useEffect, useRef, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist/build/pdf.mjs";
import pdfjsWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { supabase } from "./supabase.js";
import { exportUnpaidXlsx } from "./unpaidExport.js";
import { downloadStyledXlsx, XC, xDate } from "./xlsxStyled.js";
import { TR_UK } from "./i18n_uk.js";
import { POSITIONS, RESIDENCE, MESSENGERS, START, HEARD, CAND_STATUS, labelOf } from "./candidates.js";
import { FIN_CATS, CZ_CIT, CZ_DIVIDEND, catLabel, catKind, suggestCategory, ruleMatches, txText, fold, decodeText, parseCsv, findHeader, guessMapping, FIELDS, rowsToTxs, finYear, taxEstimate, vatOf, toCzkAmt, payrollLevies, citAdvances, CZ_SOC_FIRM, CZ_SOC_EMP, CZ_ZDR_FIRM, CZ_ZDR_EMP } from "./finance.js";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorkerUrl;

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
const MONTH_HOURS_WARN = 320; // iznad ovoga upozoravamo da su sati vjerojatno krivo upisani
const isHoursSuspicious = (h) => h > MONTH_HOURS_WARN;
const DAY_HOURS_WARN = 20; // iznad ovoga upozoravamo za jedan dan (jednokratan upis, ne mjesečni zbroj)
const isDayHoursSuspicious = (h) => h > DAY_HOURS_WARN;
const prevMonthKey = (mo) => { const [y, m] = mo.split("-").map(Number); const d = new Date(y, m - 2, 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; };
// širok raspon "normalne" satnice po valuti — samo za hvatanje očitih grešaka u tipkanju (npr. 70 umjesto 7)
const isRateSuspicious = (rate, cur) => { const r = Number(rate) || 0; if (!r) return false; return cur === "CZK" ? (r < 50 || r > 600) : (r < 2 || r > 30); };
const BRAND = "Molnar Rose";
// do 15. u mjesecu obično se obrađuje prošli mjesec (sati, isplate, fakture) — tad je on zadani
const workMonthDefault = () => { const n = new Date(); const d = n.getDate() <= 15 ? new Date(n.getFullYear(), n.getMonth() - 1, 1) : n; return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; };
function BrandMark({ size = 34 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 40 40" aria-hidden="true" style={{ flex: "0 0 auto" }}>
      <rect width="40" height="40" rx="11" fill="#0E6B4F" />
      <path d="M20 9c3.8 2.6 5.6 5.6 5.6 8.6 0 3.4-2.5 6-5.6 6s-5.6-2.6-5.6-6C14.4 14.6 16.2 11.6 20 9z" fill="#F4C9C0" />
      <path d="M20 23.6v7.4M20 27.5c-2.4-.2-4.3-1.4-5.4-3.4M20 28.6c2.2-.3 3.9-1.5 4.9-3.3" stroke="#E3F0EA" strokeWidth="1.8" strokeLinecap="round" fill="none" />
    </svg>
  );
}
const MONTHS = ["Siječanj","Veljača","Ožujak","Travanj","Svibanj","Lipanj","Srpanj","Kolovoz","Rujan","Listopad","Studeni","Prosinac"];
const MONTHS_UK = ["Січень","Лютий","Березень","Квітень","Травень","Червень","Липень","Серпень","Вересень","Жовтень","Листопад","Грудень"];
const MONTHS_SHORT_UK = ["січ","лют","бер","кві","тра","чер","лип","сер","вер","жов","лис","гру"];
// nazivi mjeseci za prikaz na ekranu (prema odabranom jeziku); MONTHS ostaje za ispise i povijest
const mName = (i) => (CUR_LANG === "UK" ? MONTHS_UK : MONTHS)[i];
const mShort = (i) => (CUR_LANG === "UK" ? MONTHS_SHORT_UK : MONTHS_SHORT)[i];
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
// "ivan  horvat" / "IVAN HORVAT" → "Ivan Horvat" — dosljedan format imena (prvo slovo svake riječi veliko)
const normalizeName = (s) => (s || "").trim().replace(/\s+/g, " ")
  .split(" ").map((w) => w ? w.charAt(0).toLocaleUpperCase("hr") + w.slice(1).toLocaleLowerCase("hr") : w).join(" ");
const looksLikeFullName = (s) => normalizeName(s).trim().split(" ").filter(Boolean).length >= 2;
// rastavi spremljeno "Prezime Ime" (ili staro "Ime Prezime") na dva polja za uređivanje —
// uzima zadnju riječ kao jedno polje, ostatak kao drugo (nagađanje za imena s više riječi)
const splitNameForEdit = (name) => {
  const words = normalizeName(name).trim().split(" ").filter(Boolean);
  if (words.length <= 1) return { lastName: words[0] || "", firstName: "" };
  return { lastName: words.slice(0, -1).join(" "), firstName: words[words.length - 1] };
};
// jednokratna migracija: pomakni zadnju riječ ("Ivan Horvat" → "Horvat Ivan") — za prebacivanje
// postojećih imena iz starog redoslijeda (ime prezime) u novi (prezime ime)
const swapNameOrder = (name) => {
  const words = normalizeName(name).trim().split(" ").filter(Boolean);
  if (words.length < 2) return normalizeName(name);
  return [words[words.length - 1], ...words.slice(0, -1)].join(" ");
};
const monthKey = (iso) => (iso || "").slice(0, 7);
const todayISO = () => new Date().toISOString().slice(0, 10);
const curMonth = () => { const n = new Date(); return `${n.getFullYear()}-${String(n.getMonth() + 1).padStart(2, "0")}`; };
const logSpan = (l) => (l.from && l.to ? `${l.from}–${l.to}` : l.monthly ? "mjesečni zbroj" : "upis sati");
// vizualno istakni zbirne mjesečne upise — to NISU pojedinačne smjene pa se ne mogu provjeriti po danu
const LogSpanLabel = ({ l }) => l.monthly
  ? <b style={{ color: S.amber }} title="Zbirni upis za cijeli mjesec — nema pojedinačnih smjena za provjeru">Σ mjesečni zbroj</b>
  : <>{logSpan(l)}</>;
const parseNum = (v) => parseFloat(String(v || "").replace(",", "."));
const TYPE_LABEL = { avans: "Avans", bonus: "Bonus", gorivo: "Gorivo", ostalo: "Ostali trošak", racun: "Na račun (banka)" };
const MZDY_NOTE_LABEL = "soc. + zdr.";

/* ---------- jezici (HR / CZ / UK) ---------- */
// zasad samo hrvatski i ukrajinski (CZ/RU skriveni dok ne budu potpuno prevedeni)
const LANGS = ["HR", "UK"];
// trenutni jezik za tr() — postavlja ga useLang pri svakom renderu
let CUR_LANG = "HR";
// tr("hrvatski tekst {0}", x) → prijevod na odabrani jezik (ili hrvatski ako prijevoda nema)
function tr(s, ...args) {
  const out = (CUR_LANG === "UK" && TR_UK[s]) || s;
  return args.length ? out.replace(/\{(\d+)\}/g, (m, i) => (args[i] ?? m)) : out;
}
const LANG_FLAG = { HR: "🇭🇷", CZ: "🇨🇿", UK: "🇺🇦", RU: "🇷🇺" };
const LANG_NAME = { HR: "Hrvatski", CZ: "Čeština", UK: "Українська", RU: "Русский" };
const DICT = {
  HR: {
    appTitle: "Evidencija rada - Firma", appSubtitle: "Radnici · Sati · Obračun", logout: "Odjava",
    tabDanas: "Danas", tabRadnici: "Radnici", tabObjekti: "Objekti", tabImenik: "Imenik", tabSati: "Sati", tabIsplate: "Isplate", tabObracun: "Obračun",
    loginTitle: "Prijava", email: "E-mail", password: "Lozinka", loginBtn: "Prijavi se", loginBusy: "Prijavljujem…",
    loginErr: "Neispravan e-mail ili lozinka.",
    noAccount: "Nemaš račun? Račune otvara isključivo poslodavac — javi mu se i dobit ćeš e-mail i lozinku.",
    workersTitle: "Radnici", addWorker: "+ Novi radnik", close: "Zatvori",
    fullName: "Ime i prezime *", lastName: "Prezime *", firstName: "Ime *", phone: "Broj telefona", rate: "Satnica", mainObject: "Glavni objekt", note: "Napomena",
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
    tabDanas: "Dnes", tabRadnici: "Pracovníci", tabObjekti: "Objekty", tabImenik: "Adresář", tabSati: "Hodiny", tabIsplate: "Výplaty", tabObracun: "Vyúčtování",
    loginTitle: "Přihlášení", email: "E-mail", password: "Heslo", loginBtn: "Přihlásit se", loginBusy: "Přihlašuji…",
    loginErr: "Nesprávný e-mail nebo heslo.",
    noAccount: "Nemáte účet? Účty zakládá pouze zaměstnavatel — ozvěte se mu a dostanete e-mail a heslo.",
    workersTitle: "Pracovníci", addWorker: "+ Nový pracovník", close: "Zavřít",
    fullName: "Jméno a příjmení *", lastName: "Příjmení *", firstName: "Jméno *", phone: "Telefonní číslo", rate: "Sazba", mainObject: "Hlavní objekt", note: "Poznámka",
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
    tabDanas: "Сьогодні", tabRadnici: "Працівники", tabObjekti: "Об'єкти", tabImenik: "Довідник", tabSati: "Години", tabIsplate: "Виплати", tabObracun: "Розрахунок",
    loginTitle: "Вхід", email: "Е-пошта", password: "Пароль", loginBtn: "Увійти", loginBusy: "Вхід…",
    loginErr: "Невірна електронна пошта або пароль.",
    noAccount: "Немає акаунту? Акаунти створює лише роботодавець — зверніться до нього і отримаєте е-пошту та пароль.",
    workersTitle: "Працівники", addWorker: "+ Новий працівник", close: "Закрити",
    fullName: "Ім'я та прізвище *", lastName: "Прізвище *", firstName: "Ім'я *", phone: "Номер телефону", rate: "Ставка", mainObject: "Основний об'єкт", note: "Примітка",
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
  RU: {
    appTitle: "Учёт работы", appSubtitle: "Работники · Часы · Расчёт", logout: "Выйти",
    tabDanas: "Сегодня", tabRadnici: "Работники", tabObjekti: "Объекты", tabImenik: "Справочник", tabSati: "Часы", tabIsplate: "Выплаты", tabObracun: "Расчёт",
    loginTitle: "Вход", email: "Эл. почта", password: "Пароль", loginBtn: "Войти", loginBusy: "Вхожу…",
    loginErr: "Неверная эл. почта или пароль.",
    noAccount: "Нет аккаунта? Аккаунты создаёт только работодатель — обратитесь к нему, и вы получите эл. почту и пароль.",
    workersTitle: "Работники", addWorker: "+ Новый работник", close: "Закрыть",
    fullName: "Имя и фамилия *", lastName: "Фамилия *", firstName: "Имя *", phone: "Номер телефона", rate: "Ставка", mainObject: "Основной объект", note: "Примечание",
    saveWorker: "Сохранить работника", noWorkersYet: "Пока нет работников. Добавьте первого кнопкой + Новый работник.",
    formerWorkers: "Бывшие работники", active: "Активен", withoutRate: "без ставки", thisMonth: "этот мес.",
    objectsTitle: "Объекты", newObjectPh: "Название нового объекта в:", addBtn: "Добавить",
    directoryTitle: "Справочник работников", searchPh: "🔍 Поиск по имени или номеру…",
    hoursTitle: "Учёт отработанных часов", object: "Объект", worker: "Работник", date: "Дата", from: "С", to: "До",
    total: "Всего", addHours: "Записать часы", recent: "Последние записи", choose: "— выбрать —",
    paymentsTitle: "Авансы, бонусы и расходы", forWorker: "👷 Для работника", forObject: "🏨 Для объекта",
    type: "Тип", amount: "Сумма (€)", save: "Сохранить",
    monthBtn: "Месяц", yearBtn: "Год", showFor: "Показать за", allObjects: "🏨 Все объекты",
    totalHours: "Всего часов", earningsBonus: "Заработок + бонусы", toPay: "К выплате", firmCosts: "Расходы фирмы",
    excelBtn: "Excel", envelopesBtn: "Конверты", cancel: "Отмена", edit: "Изменить", delete: "Удалить", back: "Назад",
  },
};
function useLang() {
  const [lang, setLang] = useState(() => { try { const l = localStorage.getItem("evidencija_lang"); return LANGS.includes(l) ? l : "HR"; } catch { return "HR"; } });
  CUR_LANG = lang;
  useEffect(() => { try { localStorage.setItem("evidencija_lang", lang); } catch {} }, [lang]);
  const t = (k) => (DICT[lang] && DICT[lang][k]) || DICT.HR[k] || k;
  return { lang, setLang, t };
}
function LangSwitch({ lang, setLang, dark, labeled }) {
  return (
    <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
      {LANGS.map((l) => (
        <button key={l} onClick={() => setLang(l)} title={LANG_NAME[l]} style={{
          background: lang === l ? (dark ? "rgba(255,255,255,.22)" : S.green) : (dark ? "rgba(255,255,255,.08)" : "#fff"),
          border: dark ? "1px solid rgba(255,255,255,.25)" : `1px solid ${S.line}`,
          borderRadius: 8, padding: "4px 7px", fontSize: 15, cursor: "pointer", lineHeight: 1 }}>
          {LANG_FLAG[l]}{labeled && <span style={{ fontSize: 12.5, fontWeight: 700, marginLeft: 5, color: lang === l ? "#fff" : S.ink }}>{l === "UK" ? tr("UA") : l}</span>}
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

/* ---------- DPH (Češka): mjesec po DUZP, rok 25. idućeg mjeseca (pomak na radni dan) ---------- */
const CZ_HOLIDAYS = ["01-01", "05-01", "05-08", "07-05", "07-06", "09-28", "10-28", "11-17", "12-24", "12-25", "12-26"];
const isCzHoliday = (iso) => {
  if (CZ_HOLIDAYS.includes(iso.slice(5))) return true;
  const e = easterISO(Number(iso.slice(0, 4)));
  return iso === addDaysISO(e, -2) || iso === addDaysISO(e, 1);
};
const dphDueDate = (taxMonth) => {
  const [y, m] = taxMonth.split("-").map(Number);
  let d = `${m === 12 ? y + 1 : y}-${String(m === 12 ? 1 : m + 1).padStart(2, "0")}-25`;
  while (isWeekend(d) || isCzHoliday(d)) d = addDaysISO(d, 1);
  return d;
};
const daysUntil = (iso) => Math.round((new Date(iso + "T12:00:00") - new Date(todayISO() + "T12:00:00")) / 86400000);
const invoiceTaxMonth = (i) => (i.duzp ? monthKey(i.duzp) : (i.period || monthKey(i.issueDate)));
const invoiceTaxBasis = (i) => (i.duzp ? "DUZP" : i.period ? "razdoblje" : "datum izdavanja");
const toCzk = (amt, cur, rate) => (cur === "EUR" ? round2((amt || 0) * rate) : (amt || 0));
function vatSummary(data, taxMonth) {
  const rate = Number(data.settings?.czk_rate) || 25;
  const out = (data.objectInvoices || []).filter((i) => invoiceTaxMonth(i) === taxMonth);
  const inn = (data.purchaseInvoices || []).filter((p) => monthKey(p.duzp) === taxMonth);
  const outVat = round2(out.reduce((sum, i) => sum + toCzk(i.vatAmount, i.vatCurrency, rate), 0));
  const inVat = round2(inn.reduce((sum, p) => sum + toCzk(p.vat, p.currency, rate), 0));
  const eurUsed = out.some((i) => i.vatAmount && i.vatCurrency === "EUR") || inn.some((p) => p.vat && p.currency === "EUR");
  return { out, inn, outVat, inVat, payable: round2(outVat - inVat), due: dphDueDate(taxMonth), eurUsed, rate };
}
function DayBadge({ iso }) {
  const h = holidayName(iso);
  if (h) return <Tag color={S.red} bg={S.redSoft}>{tr("🎉 Praznik:")}{" "}{h}</Tag>;
  if (isWeekend(iso)) return <Tag color={S.blue} bg={S.blueSoft}>{new Date(iso + "T12:00:00").getDay() === 6 ? tr("Subota") : tr("Nedjelja")}</Tag>;
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

/* ---------- dodatak na satnicu vikendom / praznikom (postavke firme) ---------- */
const payMultiplier = (settings, dateISO) => {
  if (holidayName(dateISO)) return 1 + (Number(settings?.holiday_pct) || 0) / 100;
  if (isWeekend(dateISO)) return 1 + (Number(settings?.weekend_pct) || 0) / 100;
  return 1;
};
// samo VIŠAK iznad normalne satnice za taj upis (0 ako je radni dan ili nema postavljen dodatak)
const extraPayForLog = (data, w, l) => {
  const mult = payMultiplier(data.settings, l.date);
  if (mult <= 1) return 0;
  return round2(l.hours * rateFor(data, w, l.date) * (mult - 1));
};

/* ---------- red za slanje kad nema interneta (trenutno: upis sati) ---------- */
const OFFLINE_QUEUE_KEY = "evidencija_offline_queue";
const readOfflineQueue = () => { try { return JSON.parse(localStorage.getItem(OFFLINE_QUEUE_KEY) || "[]"); } catch { return []; } };
const writeOfflineQueue = (q) => { try { localStorage.setItem(OFFLINE_QUEUE_KEY, JSON.stringify(q)); } catch {} };
const isNetworkError = (e) => (typeof navigator !== "undefined" && navigator.onLine === false)
  || /fetch|network|failed to fetch/i.test(String(e?.message || e || ""));

/* ---------- naplata objekta po poziciji (HSK, kuhinja, bar…) ---------- */
const positionBillFor = (data, objectId, position) =>
  (data.positionBilling || []).find((pb) => pb.objectId === objectId && pb.position === (position || ""));
function billRateForLog(data, object, worker, date) {
  const pos = worker?.position || "";
  const entry = pos ? positionBillFor(data, object?.id, pos) : null;
  if (entry) {
    if (entry.mode === "markup") return { rate: round2((rateFor(data, worker, date) || 0) + (entry.markup || 0)), cur: entry.currency };
    return { rate: entry.rate || 0, cur: entry.currency };
  }
  return { rate: object?.billRate || 0, cur: object?.billCur || "EUR" };
}
// brigade (po danu, bez imena) se naplaćuje hotelu po zadanoj cijeni objekta (nema pozicije/radnika)
function objRevenue(data, object, logs, brigade) {
  let revenue = 0, revenueKc = 0;
  const byPos = new Map();
  logs.forEach((l) => {
    const wk = data.workers.find((x) => x.id === l.workerId);
    const { rate, cur } = billRateForLog(data, object, wk, l.date);
    const amt = round2(l.hours * rate);
    if (cur === "CZK") revenueKc = round2(revenueKc + amt); else revenue = round2(revenue + amt);
    const key = (wk?.position || "") + "|" + cur;
    const row = byPos.get(key) || { position: wk?.position || "", cur, hours: 0, rate, amount: 0 };
    row.hours = round2(row.hours + l.hours);
    row.amount = round2(row.amount + amt);
    byPos.set(key, row);
  });
  (brigade || []).forEach((b) => {
    // posebna cijena te brigade (ako je upisana) nadjačava zadanu cijenu objekta
    const rate = b.billRate > 0 ? b.billRate : (object?.billRate || 0);
    const cur = b.billRate > 0 ? b.billCur : (object?.billCur || "EUR");
    const amt = round2(b.hours * rate);
    if (cur === "CZK") revenueKc = round2(revenueKc + amt); else revenue = round2(revenue + amt);
    const key = "__brigada__|" + cur + "|" + rate;
    const row = byPos.get(key) || { position: "Brigada (bez imena)", cur, hours: 0, rate, amount: 0 };
    row.hours = round2(row.hours + b.hours);
    row.amount = round2(row.amount + amt);
    byPos.set(key, row);
  });
  return { revenue, revenueKc, byPosition: [...byPos.values()].sort((a, b) => a.position.localeCompare(b.position, "hr")) };
}

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
    h3{font-size:12px;text-transform:uppercase;letter-spacing:.6px;color:#5C6862;margin:20px 0 0}
  </style></head><body>${bodyHtml}<script>setTimeout(function(){window.print()},350)<\/script></body></html>`);
  w.document.close();
}
// polje "oib" služi i za češki IČO: 8 znamenki = IČO (DIČ plátce DPH je CZ + IČO), 11 znamenki = hrvatski OIB
const firmIdLabel = (id) => {
  const v = String(id || "").replace(/\s+/g, "");
  if (!v) return "";
  return /^\d{8}$/.test(v) ? `IČO: ${v} · DIČ: CZ${v}` : `OIB: ${v}`;
};
const firmHead = (st) => `<div class="head"><h1>${st.company_name || "Moja firma"}</h1>
  <div class="muted">${[st.address, firmIdLabel(st.oib), st.iban ? "IBAN: " + st.iban : ""].filter(Boolean).join(" · ")}</div></div>`;
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
    ${inv.vatAmount ? `<div class="muted">Od čega PDV: <b>${money(inv.vatAmount, inv.vatCurrency || inv.currency)}</b></div>` : ""}
    ${st.iban ? `<div class="muted">Uplata na IBAN: <b>${st.iban}</b></div>` : ""}
    <div class="muted" style="margin-top:20px">Datum dospijeća: ______________________</div>`);
}

/* ---------- čitanje "Mzdy" PDF-a (češke platne liste): tekst → redovi → podaci ---------- */
async function extractPdfLines(file) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const lines = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const rows = new Map(); // zaokruženi Y → stavke
    content.items.forEach((it) => {
      const y = Math.round(it.transform[5] / 2) * 2;
      if (!rows.has(y)) rows.set(y, []);
      rows.get(y).push({ x: it.transform[4], s: it.str });
    });
    const ys = [...rows.keys()].sort((a, b) => b - a);
    ys.forEach((y) => {
      const line = rows.get(y).sort((a, b) => a.x - b.x).map((it) => it.s).join(" ").replace(/\s+/g, " ").trim();
      if (line) lines.push(line);
    });
  }
  return lines;
}
const czNum = (s) => parseFloat(String(s || "").replace(/\s/g, "").replace(",", "."));
const normName = (s) => (s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "")
  .toLowerCase().replace(/[^a-z\s]/g, " ").split(/\s+/).filter(Boolean).sort().join(" ");
const findWorkerByName = (workers, rawName) => {
  const target = normName(rawName);
  return workers.find((w) => normName(w.name) === target);
};

/* ---------- Mzdy PDF-ovi od knjigovođe: pásky, oprava, mzdové náklady, příkazy k úhradě ---------- */
// soc.-zdr. koji se odbija na koverti (fiksno prema ugovoru): puno radno vrijeme / pola / DPČ-DPP
const SOCZDR_FULL = 5500, SOCZDR_HALF = 3300, FULL_TARIF = 22400;
const socZdrFor = (p) => (/^DP[PČC]$/i.test(p.kat || "") || (!p.zdrav && !p.soc) ? 0 : (p.tarif >= FULL_TARIF ? SOCZDR_FULL : SOCZDR_HALF));
const mzdyAmount = (line) => {
  const nums = [...line.matchAll(/(\d[\d\s]*(?:[.,]\d+)?)\s*K\s*č/gi)];
  return nums.length ? czNum(nums[nums.length - 1][1]) : 0;
};
// vrsta dokumenta po sadržaju
function mzdyDocType(lines) {
  const txt = lines.slice(0, 15).join(" ");
  if (/MZDOV.{1,3}\s*N.{1,3}KLADY/i.test(txt)) return "naklady";
  if (/QR\s*platby/i.test(txt) || lines.some((l) => /z\s*pr.{1,3}va\s*p.{1,3}jemci/i.test(l))) return "prikazy";
  return "pasky";
}
// výplatní pásky (i "oprava") → po radniku
function parseMzdyPasky(lines) {
  const out = [];
  let cur = null;
  const push = () => { if (cur && cur.name) out.push(cur); cur = null; };
  for (const line of lines) {
    const m = line.match(/(\d{2,6})\s*\S{0,3}\s*([A-ZÁ-Ž][A-Za-zÁ-Žá-ž]+(?:\s+[A-Za-zÁ-Žá-ž]+)+?)\s*\S{0,3}\s*(\d{6}\/\d{3,4})\s*\S{0,3}\s*(\S{2,4})\s*\S{0,3}\s*(\d{4})\s*\/\s*(\d{2})/);
    if (m) { push(); cur = { name: m[2].trim(), kat: m[4].toUpperCase(), month: `${m[5]}-${m[6]}`, tarif: 0, zdrav: 0, soc: 0, pojist: 0, dobirka: 0, hruba: 0 }; continue; }
    if (!cur) continue;
    const t = line.match(/M.{1,3}s.{1,3}n.{1,3}\s*tarif\s+(\d[\d\s]*)\s*K\s*č/i);
    if (t) cur.tarif = czNum(t[1]);
    if (/Zdrav\.?\s*poj/i.test(line)) cur.zdrav = mzdyAmount(line);
    else if (/Soc\.?\s*poj/i.test(line)) cur.soc = mzdyAmount(line);
    else if (/Pojistn.{0,3}zam.{0,4}stnavatel/i.test(line)) cur.pojist = mzdyAmount(line);
    else if (/Hrub.{1,2}\s*mzda/i.test(line)) cur.hruba = mzdyAmount(line);
    if (/Dob.{1,2}rka/i.test(line)) cur.dobirka = mzdyAmount(line);
  }
  push();
  return out;
}
// mzdové náklady → { prezime: dobírka } za provjeru
function parseMzdyNaklady(lines) {
  const out = [];
  for (const line of lines) {
    const m = line.match(/^([A-ZÁ-Ž]{2,}[A-ZÁ-Ž\-]*)\s+([A-Za-zÁ-Žá-ž]+)\s+(\d[\d\s]*)$/);
    if (!m) continue;
    const nums = m[3].trim().split(/\s+/).map(Number);
    if (nums.length < 2) continue;
    out.push({ surname: m[1], first: m[2], hruba: nums[0], dobirka: nums[nums.length - 1] });
  }
  return out;
}
// příkazy k úhradě (QR platby) → uplate
function parseMzdyPrikazy(lines) {
  const out = [];
  let cur = null;
  for (const line of lines) {
    if (/p.{1,3}jemce\s*:/i.test(line) && !/zpr/i.test(line)) { if (cur) out.push(cur); cur = { to: line.split(":").slice(1).join(":").trim(), amount: 0, msg: "" }; continue; }
    if (!cur) continue;
    const a = line.match(/.{1,3}stka\s*:\s*([\d\s]+(?:[.,]\d+)?)/i);
    if (a) cur.amount = czNum(a[1]);
    const z = line.match(/zpr.{1,3}va\s*p.{1,3}jemci\s*:\s*(.*)$/i);
    if (z) cur.msg = z[1].trim();
  }
  if (cur) out.push(cur);
  return out.filter((p) => p.amount > 0).map((p) => ({
    ...p, to: p.to || (/da.{1,2}/i.test(p.msg) ? "Finanční úřad" : ""),
    label: /zdrav/i.test(p.msg) ? "zdravotní pojištění" : /soci/i.test(p.msg) ? "sociální pojištění" : /da/i.test(p.msg) ? "daň ze mzdy" : p.msg,
  }));
}
// sličnost imena: isto prezime + ime s najviše 2 razlike (Ievgen ≈ Jevgen)
const editDist = (a, b) => {
  const d = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
  for (let j = 1; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++)
    d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  return d[a.length][b.length];
};
const plainTokens = (s) => (s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z\s]/g, " ").split(/\s+/).filter(Boolean);
function matchWorkerFuzzy(workers, rawName) {
  const exact = findWorkerByName(workers, rawName);
  if (exact) return exact;
  const [sur, ...rest] = plainTokens(rawName);
  const first = rest.join(" ");
  const cands = workers.filter((w) => {
    const t = plainTokens(w.name);
    if (!t.includes(sur) && !t.some((x) => editDist(x, sur) <= 1)) return false;
    const others = t.filter((x) => x !== sur).join(" ");
    return !first || !others || editDist(others, first) <= 2;
  });
  const active = cands.filter((w) => !w.archived);
  return (active.length === 1 ? active[0] : cands.length === 1 ? cands[0] : null);
}

/* ---------- čitanje izlaznih faktura (PDF) — izvlači iznos, PDV, kupca, broj, datum ---------- */
// kao extractPdfLines, ali čuva x-poziciju svake riječi po retku (treba za dvostupačni
// "Dodavatel / Odběratel" dio fakture) I vraća stranice ODVOJENO — jedan PDF zna sadržavati
// više faktura zaredom (npr. izvoz iz Vyfakturuj.cz), svaka počinje na svojoj stranici
async function extractPdfPages(file) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const rows = new Map();
    content.items.forEach((it) => {
      const y = Math.round(it.transform[5] / 2) * 2;
      if (!rows.has(y)) rows.set(y, []);
      rows.get(y).push({ x: it.transform[4], s: it.str });
    });
    const ys = [...rows.keys()].sort((a, b) => b - a);
    const lines = ys.map((y) => rows.get(y).sort((a, b) => a.x - b.x).filter((it) => it.s.trim())).filter((r) => r.length);
    pages.push(lines);
  }
  return pages; // [ [row,row,...], [row,row,...], ... ] — jedan niz redaka po stranici
}
// grupira stranice u pojedinačne fakture: nova faktura počinje na svakoj stranici koja ima "Faktura č. ..."
// (stranice bez tog obilježja se nadovezuju na prethodnu fakturu — npr. duga specifikacija na 2 stranice)
function splitPdfInvoices(pages) {
  const chunks = [];
  pages.forEach((page) => {
    const hasMarker = page.some((row) => /Faktura\s*č\.?\s*[A-Za-z0-9/\-]+/i.test(row.map((it) => it.s).join(" ")));
    if (hasMarker || chunks.length === 0) chunks.push([page]);
    else chunks[chunks.length - 1].push(page);
  });
  return chunks.map((pageGroup) => pageGroup.flat());
}
// OCR (čitanje teksta sa slike) — za skenirane/fotografirane PDF-ove koji nemaju pravi tekst sloj.
// Renderira svaku stranicu u sliku pa pušta tesseract.js (učitava se tek kad zatreba, ne usporava app inače).
// Vraća podatke u ISTOM obliku kao extractPdfPages, da se mogu koristiti isti splitPdfInvoices/parseInvoicePdf.
async function ocrPdfPages(file, onProgress) {
  const { createWorker } = await import("tesseract.js");
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const worker = await createWorker("ces");
  const pages = [];
  try {
    for (let p = 1; p <= pdf.numPages; p++) {
      if (onProgress) onProgress(p, pdf.numPages);
      const page = await pdf.getPage(p);
      const viewport = page.getViewport({ scale: 2.5 });
      const canvas = document.createElement("canvas");
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
      const { data } = await worker.recognize(canvas);
      const words = (data.words || []).filter((w) => w.text && w.text.trim());
      const rows = new Map();
      words.forEach((w) => {
        const cy = Math.round(((w.bbox.y0 + w.bbox.y1) / 2) / 12) * 12;
        if (!rows.has(cy)) rows.set(cy, []);
        rows.get(cy).push({ x: w.bbox.x0, s: w.text });
      });
      const ys = [...rows.keys()].sort((a, b) => a - b);
      const lines = ys.map((y) => rows.get(y).sort((a, b) => a.x - b.x)).filter((r) => r.length);
      pages.push(lines);
    }
  } finally {
    await worker.terminate();
  }
  return pages;
}
const MONTHS_CZ = { leden: 1, únor: 2, brezen: 3, březen: 3, duben: 4, květen: 5, kveten: 5, červen: 6, cerven: 6,
  červenec: 7, cervenec: 7, srpen: 8, září: 9, zari: 9, říjen: 10, rijen: 10, listopad: 11, prosinec: 12 };
function parseInvoicePdf(lineItems) {
  const joined = lineItems.map((row) => row.map((it) => it.s).join(" ").replace(/\s+/g, " ").trim());
  const fullText = joined.join("\n");
  // sve riječi zaredom u jedan jedini niz, bez obzira na to kako su se retci grupirali po Y-poziciji
  // (otporno na to ako je "Celkem"/"DPH"/iznos slučajno završio u susjednom, a ne istom retku)
  const flatText = lineItems.flat().map((it) => it.s).join(" ").replace(/\s+/g, " ").trim();
  // jedan broj u češkom formatu: "613 060,00" ili "613060.00" ili i bez decimala "613060"
  const NUM = "\\d[\\d\\s]*(?:[.,]\\d{1,2})?";

  const numM = fullText.match(/Faktura\s*č\.?\s*([A-Za-z0-9/\-]+)/i);
  const number = numM ? numM[1] : "";

  const dateM = fullText.match(/Datum vystaven[ií]:?\s*(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})/i);
  const issueDate = dateM ? `${dateM[3]}-${dateM[2].padStart(2, "0")}-${dateM[1].padStart(2, "0")}` : "";
  // DUZP = datum uskutečnění zdanitelného plnění — po njemu ide DPH u mjesec (ne po datumu izdavanja)
  const duzpM = fullText.match(/(?:DUZP|zdaniteln\S*\s+plněn\S*|zdan\.\s*plněn\S*)\s*:?\s*(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})/i)
    || flatText.match(/(?:DUZP|zdaniteln\S*\s+plněn\S*|zdan\.\s*plněn\S*)\s*:?\s*(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})/i);
  const duzp = duzpM ? `${duzpM[3]}-${duzpM[2].padStart(2, "0")}-${duzpM[1].padStart(2, "0")}` : "";

  let amount = 0, vatAmount = 0;
  // zbirni red "Sazba | Základ | DPH | Včetně DPH", npr. "21 % 613 060,00 Kč 128 742,60 Kč 741 802,60 Kč" —
  // NE oslanja se na to da je "Kč" ispravno pročitan (ponekad pdfjs razdvoji ili izgubi taj simbol),
  // nego samo traži tri broja s decimalama jedan iza drugog nakon postotka
  const vatRe = new RegExp(`(\\d{1,2})\\s*%\\D{0,10}(${NUM})\\D{0,10}(${NUM})\\D{0,10}(${NUM})`);
  const vatRowM = flatText.match(vatRe) || fullText.match(vatRe);
  if (vatRowM) {
    vatAmount = czNum(vatRowM[3]);
    amount = czNum(vatRowM[4]); // "Včetně DPH" = ukupan iznos s PDV-om
  }
  // "Celkem k úhradě: ..." — ako postoji, to je pouzdaniji izvor ukupnog iznosa, prebije gornje
  const totalRe = new RegExp(`Celkem k úhrad[ěe]:?\\D{0,6}(${NUM})`, "i");
  const totalM = flatText.match(totalRe) || fullText.match(totalRe);
  if (totalM) amount = czNum(totalM[1]);
  if (!amount) {
    // posljednji fallback: zadnji broj s dvije decimale u cijelom dokumentu (obično je to ukupan iznos)
    const allNums = [...flatText.matchAll(/\d[\d\s]*[.,]\d{2}/g)];
    if (allNums.length) amount = czNum(allNums[allNums.length - 1][0]);
  }

  let period = "";
  const perM = fullText.match(/za m[ěe]s[ií]c\s+([A-Za-zÁ-Žá-ž]+)\s+(\d{4})/i);
  if (perM) {
    const mo = MONTHS_CZ[perM[1].toLowerCase()];
    if (mo) period = `${perM[2]}-${String(mo).padStart(2, "0")}`;
  }
  if (!period && issueDate) period = issueDate.slice(0, 7);

  // "Dodavatel" i "Odběratel" stoje jedan kraj drugog u istom retku (zaglavlje dva stupca);
  // ime kupca je u sljedećem retku, na x-poziciji poravnatoj s "Odběratel" (desni stupac)
  let customerName = "";
  for (let i = 0; i < lineItems.length - 1; i++) {
    const row = lineItems[i];
    const odb = row.find((it) => /^Odběratel$/i.test(it.s.trim()));
    const dod = row.find((it) => /^Dodavatel$/i.test(it.s.trim()));
    if (odb && dod) {
      const nameRow = lineItems[i + 1] || [];
      const rightTokens = nameRow.filter((it) => it.x >= odb.x - 5);
      customerName = rightTokens.map((it) => it.s).join(" ").replace(/\s+/g, " ").trim();
      break;
    }
  }

  return { number, issueDate, duzp, amount, vatAmount, period, customerName, currency: "CZK", rawText: flatText };
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

  let billing = [], audit = [], trash = [], members = [], invoicePayments = [], settings = {}, objectInvoices = [], payrollNotes = [], positionBilling = [], reminders = [], loginLog = [], errorLog = [], commissionRates = [], commissionPayouts = [], brigadeCosts = [], purchaseInvoices = [], monthClose = [], cashAdvances = [];
  let bankMap = {};
  // nove prijave kandidata (RLS: vidi samo admin ili regruter, ostali dobiju 0)
  let newCandidates = 0, staleCandidates = 0;
  {
    const dayAgo = new Date(Date.now() - 24 * 3600000).toISOString();
    const [cn, cs] = await Promise.all([
      supabase.from("candidates").select("id", { count: "exact", head: true }).eq("status", "novi"),
      supabase.from("candidates").select("id", { count: "exact", head: true }).eq("status", "novi").lt("created_at", dayAgo),
    ]);
    newCandidates = cn.error ? 0 : (cn.count || 0);
    staleCandidates = cs.error ? 0 : (cs.count || 0);
  }
  if (isAdmin) {
    const wb = await supabase.from("worker_bank").select("worker_id, bank_account");
    if (!wb.error) (wb.data || []).forEach((r) => { bankMap[r.worker_id] = r.bank_account || ""; });
    const [b, a, tw, tl, tp, om, ip, st, oi, pn, pb, rm, ll, el, cr, cp, bg, pi, mc, ca] = await Promise.all([
      supabase.from("object_billing").select("*"),
      supabase.from("audit_log").select("*").order("at", { ascending: false }).limit(80),
      supabase.from("workers").select("*").not("deleted_at", "is", null),
      supabase.from("work_logs").select("*").not("deleted_at", "is", null).order("deleted_at", { ascending: false }).limit(50),
      supabase.from("payments").select("*").not("deleted_at", "is", null).order("deleted_at", { ascending: false }).limit(50),
      supabase.from("object_members").select("*"),
      live(supabase.from("invoice_payments").select("*")),
      supabase.from("settings").select("*").eq("id", 1).maybeSingle(),
      live(supabase.from("object_invoices").select("*")).order("issue_date", { ascending: false }),
      live(supabase.from("payroll_notes").select("*")),
      supabase.from("position_billing").select("*"),
      live(supabase.from("reminders").select("*")),
      supabase.from("login_log").select("*").order("at", { ascending: false }).limit(80),
      supabase.from("error_log").select("*").order("at", { ascending: false }).limit(50),
      supabase.from("commission_rates").select("*"),
      live(supabase.from("commission_payouts").select("*")),
      live(supabase.from("brigade_costs").select("*")),
      live(supabase.from("purchase_invoices").select("*")).order("duzp", { ascending: false }),
      supabase.from("month_close").select("*"),
      live(supabase.from("cash_advances").select("*")),
    ]);
    billing = b.data || []; audit = a.data || []; members = om.data || [];
    invoicePayments = ip.data || []; settings = st.data || {}; objectInvoices = oi.data || []; payrollNotes = pn.data || [];
    positionBilling = pb.data || []; reminders = rm.data || []; loginLog = ll.data || []; errorLog = el.data || []; commissionRates = cr.data || []; commissionPayouts = cp.data || []; brigadeCosts = bg.data || []; purchaseInvoices = pi.data || []; monthClose = mc.data || []; cashAdvances = ca.data || [];
    trash = [
      ...(tw.data || []).map((r) => ({ table: "workers", row: r, label: `Radnik: ${r.name}` })),
      ...(tl.data || []).map((r) => ({ table: "work_logs", row: r, label: `Sati: ${fmtH(r.hours)} (${fmtDate(r.work_date)})` })),
      ...(tp.data || []).map((r) => ({ table: "payments", row: r, label: `${TYPE_LABEL[r.type] || r.type}: ${eur(r.amount)} (${fmtDate(r.pay_date)})` })),
    ].sort((a, b) => (b.row.deleted_at || "").localeCompare(a.row.deleted_at || ""));
  }
  const billMap = Object.fromEntries(billing.map((b) => [b.object_id, { rate: Number(b.bill_rate) || 0, cur: b.bill_currency === "CZK" ? "CZK" : "EUR", budget: Number(b.monthly_budget) || 0 }]));

  return {
    workers: (workers.data || []).map((w) => ({
      id: w.id, name: w.name, phone: w.phone || "", rate: Number(w.base_rate) || 0,
      objectId: w.object_id || "", note: w.note || "", archived: !!w.archived, archivedDate: w.archived_date || "",
      permitExpiry: w.permit_expiry || "", contractExpiry: w.contract_expiry || "", position: w.position || "", bankAccount: bankMap[w.id] || "",
      rateCurrency: w.rate_currency === "CZK" ? "CZK" : "EUR",
    })),
    objects: (objects.data || []).map((o) => ({
      id: o.id, name: o.name, billRate: billMap[o.id]?.rate || 0, billCur: billMap[o.id]?.cur || countryCur(o.country || "HR"), country: o.country || "HR",
      archived: !!o.archived, archivedDate: o.archived_date || "", monthlyBudget: billMap[o.id]?.budget || 0,
      invoiceCustomerName: o.invoice_customer_name || "",
    })),
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
    payouts: (payouts.data || []).map((p) => ({
      id: p.id, workerId: p.worker_id, month: p.month, amount: Number(p.amount), amountKc: Number(p.amount_czk) || 0, paidAt: p.paid_at,
      approved: !!p.approved, approvedBy: p.approved_by || "", approvedAt: p.approved_at || "",
    })),
    invoicePayments: invoicePayments.map((p) => ({ id: p.id, objectId: p.object_id, month: p.month, amount: Number(p.amount), date: p.pay_date, note: p.note || "" })),
    objectInvoices: objectInvoices.map((i) => ({
      id: i.id, objectId: i.object_id, number: i.number, period: i.period, issueDate: i.issue_date, duzp: i.duzp || "",
      hours: Number(i.hours) || 0, rate: Number(i.rate) || 0, amount: Number(i.amount) || 0,
      currency: i.currency === "CZK" ? "CZK" : "EUR", note: i.note || "",
      vatAmount: Number(i.vat_amount) || 0, vatCurrency: i.vat_currency === "CZK" ? "CZK" : "EUR", filePath: i.file_path || "",
    })),
    payrollNotes: payrollNotes.map((n) => ({
      id: n.id, workerId: n.worker_id, month: n.month, amount: Number(n.amount) || 0,
      currency: n.currency === "EUR" ? "EUR" : "CZK", sourceCzk: Number(n.source_czk) || 0, note: n.note || "",
    })),
    positionBilling: positionBilling.map((pb) => ({
      objectId: pb.object_id, position: pb.position || "", mode: pb.mode === "markup" ? "markup" : "fixed",
      rate: Number(pb.rate) || 0, markup: Number(pb.markup) || 0, currency: pb.currency === "CZK" ? "CZK" : "EUR",
    })),
    reminders: reminders.map((r) => ({ id: r.id, text: r.text, dueDate: r.due_date || "", done: !!r.done })),
    commissionRates: commissionRates.map((c) => ({ objectId: c.object_id, userId: c.user_id, rate: Number(c.rate) || 0, currency: c.currency === "CZK" ? "CZK" : "EUR" })),
    commissionPayouts: commissionPayouts.map((c) => ({
      id: c.id, userId: c.user_id, payDate: c.pay_date, amount: Number(c.amount) || 0,
      currency: c.currency === "CZK" ? "CZK" : "EUR", note: c.note || "",
    })),
    brigadeCosts: brigadeCosts.map((b) => ({
      id: b.id, objectId: b.object_id, date: b.work_date, hours: Number(b.hours) || 0,
      amount: Number(b.amount) || 0, currency: b.currency === "EUR" ? "EUR" : "CZK", note: b.note || "",
      billRate: Number(b.bill_rate) || 0, billCur: b.bill_currency === "CZK" ? "CZK" : "EUR",
      people: Number(b.people) || 1, names: b.names || "", paidBy: b.paid_by || null,
    })),
    // gotovina koju admin da koordinatoru (given) ili mu ovaj vrati višak (returned)
    cashAdvances: cashAdvances.map((c) => ({
      id: c.id, profileId: c.profile_id, kind: c.kind === "returned" ? "returned" : "given", date: c.pay_date,
      amount: Number(c.amount) || 0, currency: c.currency === "EUR" ? "EUR" : "CZK", note: c.note || "",
    })),
    purchaseInvoices: purchaseInvoices.map((p) => ({
      id: p.id, supplier: p.supplier || "", number: p.number || "", duzp: p.duzp, base: Number(p.base_amount) || 0,
      vat: Number(p.vat_amount) || 0, currency: p.currency === "EUR" ? "EUR" : "CZK", note: p.note || "",
    })),
    monthClose: monthClose.map((c) => ({ id: c.id, month: c.month, objectId: c.object_id || null, step: c.step, doneAt: c.done_at })),
    loginLog: loginLog.map((l) => ({ id: l.id, userName: l.user_name || "", at: l.at, userAgent: l.user_agent || "" })),
    errorLog: errorLog.map((e) => ({ id: e.id, userName: e.user_name || "", message: e.message, url: e.url || "", at: e.at })),
    settings: settings || {},
    newCandidates, staleCandidates,
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
  const [offlineQueue, setOfflineQueue] = useState(() => readOfflineQueue());
  const [tab, setTab] = useState("danas");
  const [openWorker, setOpenWorker] = useState(null);
  const [openObject, setOpenObject] = useState(null);
  const [openProfile, setOpenProfile] = useState(null);
  const [adminPanel, setAdminPanel] = useState("");

  const admin = profile?.role === "admin";
  const loginLogged = useRef(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: sub } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  const reload = useCallback(async (prof) => {
    const p = prof || profile;
    if (!p) return;
    try { setData(await fetchAll(p.role === "admin")); setErr(""); }
    catch (e) { setErr(tr("Greška pri učitavanju: ") + (e.message || e)); }
  }, [profile]);

  const queuePush = (type, payload) => {
    const q = readOfflineQueue();
    q.push({ id: Date.now() + "_" + Math.random().toString(36).slice(2), type, payload, createdAt: Date.now() });
    writeOfflineQueue(q);
    setOfflineQueue(q);
  };
  const insertQueued = async (item) => {
    const table = item.type === "work_log" ? "work_logs" : item.type;
    const { error } = await supabase.from(table).insert(item.payload);
    if (error) throw error;
  };
  const flushOfflineQueue = useCallback(async () => {
    const q = readOfflineQueue();
    if (!q.length || typeof navigator !== "undefined" && navigator.onLine === false) return;
    const remaining = [];
    for (const item of q) {
      try { await insertQueued(item); } catch { remaining.push(item); }
    }
    writeOfflineQueue(remaining);
    setOfflineQueue(remaining);
    if (remaining.length !== q.length) await reload();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [reload]);
  useEffect(() => {
    flushOfflineQueue();
    window.addEventListener("online", flushOfflineQueue);
    return () => window.removeEventListener("online", flushOfflineQueue);
  }, [flushOfflineQueue]);

  // osnovno bilježenje grešaka aplikacije (bez vanjskog servisa) — vidljivo adminu u 🐛 Greške
  useEffect(() => {
    const report = (message, stack) => {
      if (!session) return;
      supabase.from("error_log").insert({
        user_id: session.user.id, user_name: profile?.name || "", message: String(message || "").slice(0, 500),
        stack: String(stack || "").slice(0, 2000), url: window.location.href,
      }).then(() => {}).catch(() => {});
    };
    const onError = (e) => report(e.message, e.error?.stack);
    const onRejection = (e) => report("Unhandled promise rejection: " + (e.reason?.message || e.reason), e.reason?.stack);
    window.addEventListener("error", onError);
    window.addEventListener("unhandledrejection", onRejection);
    return () => { window.removeEventListener("error", onError); window.removeEventListener("unhandledrejection", onRejection); };
  }, [session, profile]);

  useEffect(() => {
    if (!session) { setProfile(null); setData(null); return; }
    (async () => {
      const { data: p, error } = await supabase.from("profiles").select("*").eq("id", session.user.id).single();
      if (error) { setErr(tr("Ne mogu učitati profil: ") + error.message); return; }
      setProfile(p);
      await reload(p);
      if (!loginLogged.current) {
        loginLogged.current = true;
        supabase.from("login_log").insert({ user_id: session.user.id, user_name: p.name || "", user_agent: navigator.userAgent }).then(() => {});
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [session]);

  useEffect(() => {
    const onFocus = () => { if (profile) reload(); };
    window.addEventListener("focus", onFocus);
    return () => window.removeEventListener("focus", onFocus);
  }, [profile, reload]);

  // lokalna obavijest (dok je app otvorena/pokrenuta) — jednom dnevno: istek dokumenata + (admin) neisplaćeni prošli mjesec
  useEffect(() => {
    if (!data || typeof Notification === "undefined" || Notification.permission !== "granted") return;
    const key = "evidencija_notif_" + todayISO();
    try { if (localStorage.getItem(key)) return; } catch { return; }
    const parts = [];
    const warns = expiryWarnings(data.workers, data.settings.expiry_warn_days);
    if (warns.length) {
      const expired = warns.filter((w) => w.past).length;
      const soon = warns.length - expired;
      parts.push([expired ? `${expired} dokument(a) isteklo` : "", soon ? `${soon} uskoro ističe` : ""].filter(Boolean).join(", "));
    }
    if (admin) {
      const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1);
      const prevMo = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      const withHours = new Set(data.logs.filter((l) => monthKey(l.date) === prevMo).map((l) => l.workerId));
      const unpaid = [...withHours].filter((wid) => !paidFor(data, wid, prevMo)?.approved).length;
      if (unpaid > 0) parts.push(`${unpaid} radnika još nije odobreno za isplatu za ${MONTHS[Number(prevMo.slice(5, 7)) - 1]}`);
      const paydayDay = Number(data.settings.payday_day) || 5;
      const reminderDays = Number(data.settings.payday_reminder_days) || 3;
      const todayDay = new Date().getDate();
      const daysLeft = paydayDay - todayDay;
      if (unpaid > 0 && daysLeft >= 0 && daysLeft <= reminderDays) {
        parts.push(daysLeft === 0 ? `💰 Danas je isplata, a još ima ${unpaid} neodobrenih` : `💰 Isplata za ${daysLeft} ${daysLeft === 1 ? "dan" : "dana"} — još ima ${unpaid} neodobrenih`);
      }
      const dueReminders = (data.reminders || []).filter((r) => !r.done && r.dueDate && r.dueDate <= todayISO());
      if (dueReminders.length) parts.push(`${dueReminders.length} podsjetnik(a) dospjelo`);
    }
    if (!parts.length) return;
    try {
      new Notification("Evidencija rada", { body: parts.join(" · "), icon: "/icon-192.png" });
      localStorage.setItem(key, "1");
    } catch {}
  }, [data, admin]);

  const act = async (fn, auditText) => {
    setBusy(true);
    try {
      await fn();
      if (auditText) await supabase.from("audit_log").insert({ user_id: session.user.id, user_name: profile?.name || "", action: auditText });
      await reload();
      setErr("");
      return true;
    } catch (e) { setErr(tr("Greška: ") + (e.message || e)); return false; }
    finally { setBusy(false); }
  };
  const ins = (table, row) => supabase.from(table).insert(row).then(({ error }) => { if (error) throw error; });
  const upd = (table, id, patch) => supabase.from(table).update(patch).eq("id", id).then(({ error }) => { if (error) throw error; });
  const softDel = (table, id) => upd(table, id, { deleted_at: new Date().toISOString() });
  const saveBank = (workerId, acc) => supabase.from("worker_bank").upsert({ worker_id: workerId, bank_account: (acc || "").trim(), updated_by: session.user.id, updated_at: new Date().toISOString() }).then(({ error }) => { if (error) throw error; });
  const guardPaid = (workerId, dateOrMonth) => {
    const mo = dateOrMonth.length === 7 ? dateOrMonth : monthKey(dateOrMonth);
    const p = paidFor(data, workerId, mo);
    if (p && p.approved) throw new Error(tr("Mjesec {0} je ODOBREN i zaključan za tog radnika. Admin ga može otključati u Obračunu.", mo));
  };

  const api = {
    admin, recruiter: admin || !!profile?.can_recruit, myName: (profile?.name || "").split(" ")[0], uid: () => session?.user?.id, settings: data?.settings || {}, t, lang, setLang,
    run: (fn, auditText) => act(fn, auditText),
    addWorker: (f) => act(async () => {
      const name = normalizeName(f.name);
      if (!name) throw new Error(tr("Ime radnika ne smije biti prazno."));
      const { data: w, error } = await supabase.from("workers")
        .insert({ name, phone: f.phone, base_rate: parseNum(f.rate) || 0, rate_currency: f.rateCur || "EUR", object_id: f.objectId || null, note: f.note,
          permit_expiry: f.permitExpiry || null, contract_expiry: f.contractExpiry || null, position: f.position || "", created_by: session.user.id })
        .select().single();
      if (error) throw error;
      if (f.objectId) await ins("assignments", { worker_id: w.id, object_id: f.objectId, from_date: todayISO(), created_by: session.user.id });
      if (admin && (f.bankAccount || "").trim()) await saveBank(w.id, f.bankAccount);
    }, `Dodao radnika: ${normalizeName(f.name)}`),
    addWorkersBulk: (rows) => act(async () => {
      const clean = rows.map((r) => ({ ...r, name: normalizeName(r.name) })).filter((r) => r.name);
      if (!clean.length) throw new Error(tr("Nema redaka s imenom za uvoz."));
      const payload = clean.map((r) => ({
        name: r.name, phone: r.phone || "", base_rate: parseNum(r.rate) || 0, rate_currency: r.rateCur || "EUR",
        object_id: r.objectId || null, position: r.position || "", note: r.note || "",
        permit_expiry: r.permitExpiry || null, contract_expiry: r.contractExpiry || null, created_by: session.user.id,
      }));
      const { error } = await supabase.from("workers").insert(payload);
      if (error) throw error;
    }, `Uvezao ${rows.length} radnika iz Excela`),
    updWorker: (id, f, name) => {
      const newName = normalizeName(f.name);
      if (!newName) { setErr(tr("Greška: Ime radnika ne smije biti prazno.")); return Promise.resolve(false); }
      const old = data.workers.find((w) => w.id === id);
      const objName = (oid) => data.objects.find((o) => o.id === oid)?.name || "—";
      const diffs = [];
      if (old) {
        if (old.name !== newName) diffs.push(`ime "${old.name}"→"${newName}"`);
        if ((old.phone || "") !== (f.phone || "")) diffs.push(`telefon "${old.phone || "—"}"→"${f.phone || "—"}"`);
        if (round2(old.rate || 0) !== round2(parseNum(f.rate) || 0)) diffs.push(`satnica ${old.rate || 0}→${parseNum(f.rate) || 0}`);
        if ((old.objectId || "") !== (f.objectId || "")) diffs.push(`objekt ${objName(old.objectId)}→${objName(f.objectId)}`);
        if ((old.position || "") !== (f.position || "")) diffs.push(`pozicija "${old.position || "—"}"→"${f.position || "—"}"`);
        if ((old.permitExpiry || "") !== (f.permitExpiry || "")) diffs.push(`istek dozvole ${old.permitExpiry || "—"}→${f.permitExpiry || "—"}`);
        if ((old.contractExpiry || "") !== (f.contractExpiry || "")) diffs.push(`istek ugovora ${old.contractExpiry || "—"}→${f.contractExpiry || "—"}`);
        if ((old.bankAccount || "") !== (f.bankAccount || "").trim()) diffs.push("broj računa promijenjen");
      }
      return act(async () => { await upd("workers", id, {
        name: newName, phone: f.phone, base_rate: parseNum(f.rate) || 0, rate_currency: f.rateCur || "EUR", object_id: f.objectId || null, note: f.note,
        permit_expiry: f.permitExpiry || null, contract_expiry: f.contractExpiry || null, position: f.position || "",
      }); if (admin && (old?.bankAccount || "") !== (f.bankAccount || "").trim()) await saveBank(id, f.bankAccount); }, `Uredio podatke radnika: ${name}${diffs.length ? " (" + diffs.join(", ") + ")" : ""}`);
    },
    setPosition: (w, position) => act(() => upd("workers", w.id, { position }),
      `Postavio poziciju radniku ${w.name}: ${position || "—"}`),
    delWorker: (w) => act(() => softDel("workers", w.id), `Obrisao radnika: ${w.name}`),
    archiveWorker: (w, on) => act(() => upd("workers", w.id, {
      archived: on, archived_date: on ? todayISO() : w.archivedDate || null, object_id: on ? null : (w.objectId || null),
    }), on ? `Označio da je ${w.name} završio s radom` : `Vratio ${w.name} u aktivne`),
    archiveWorkersBulk: (workerIds) => act(async () => {
      for (const wid of workerIds) await upd("workers", wid, { archived: true, archived_date: todayISO(), object_id: null });
    }, `Grupno označio ${workerIds.length} radnika kao završene`),
    swapAllNameOrders: (workers) => act(async () => {
      for (const w of workers) await upd("workers", w.id, { name: swapNameOrder(w.name) });
      const { error } = await supabase.from("settings").upsert({ id: 1, ...data.settings, name_order_migrated: true });
      if (error) throw error;
    }, `Jednokratno zamijenio redoslijed ime/prezime → prezime/ime za ${workers.length} radnika`),
    transfer: (w, objectId, from, targetName) => act(async () => {
      await upd("workers", w.id, { object_id: objectId || null });
      await ins("assignments", { worker_id: w.id, object_id: objectId || null, from_date: from, created_by: session.user.id });
    }, `Prebacio ${w.name} na: ${targetName} (od ${fmtDate(from)})`),
    transferBulk: (workerIds, objectId, targetName, from) => act(async () => {
      for (const wid of workerIds) {
        await upd("workers", wid, { object_id: objectId || null });
        await ins("assignments", { worker_id: wid, object_id: objectId || null, from_date: from, created_by: session.user.id });
      }
    }, `Grupno prebacio ${workerIds.length} radnika na: ${targetName} (od ${fmtDate(from)})`),
    delAssignment: (a, wName) => act(() => softDel("assignments", a.id), `Obrisao premještaj radnika ${wName}`),
    changeRate: (w, rate, from) => act(() => ins("rate_changes", { worker_id: w.id, rate, from_date: from, created_by: session.user.id }),
      `Promijenio satnicu: ${w.name} → ${eur(rate)}/h od ${fmtDate(from)}`),
    delRateChange: (r, wName) => act(() => softDel("rate_changes", r.id), `Obrisao promjenu satnice radnika ${wName}`),
    addObject: async (name, country = "HR") => {
      const { data: o, error } = await supabase.from("objects").insert({ name, country, created_by: session.user.id }).select().single();
      if (error) { setErr(tr("Greška: ") + error.message); return null; }
      await supabase.from("audit_log").insert({ user_id: session.user.id, user_name: profile?.name || "", action: `Dodao objekt: ${name} (${country})` });
      await reload();
      return o.id;
    },
    setObjectCountry: (o, country) => act(() => upd("objects", o.id, { country }),
      `Premjestio objekt ${o.name} u: ${COUNTRY_NAME[country] || country}`),
    setInvoiceCustomerName: (o, name) => act(() => upd("objects", o.id, { invoice_customer_name: name }),
      `Postavio naziv kupca na fakturi za ${o.name}: "${name}" (za prepoznavanje kod uvoza PDF faktura)`),
    delObject: (o) => act(() => softDel("objects", o.id), `Obrisao objekt: ${o.name}`),
    archiveObject: (o, on) => act(() => upd("objects", o.id, { archived: on, archived_date: on ? todayISO() : null }),
      on ? `Zatvorio objekt: ${o.name}` : `Ponovno otvorio objekt: ${o.name}`),
    setBillRate: (o, rate, cur) => act(() => supabase.from("object_billing").upsert({ object_id: o.id, bill_rate: rate, bill_currency: cur || "EUR" })
      .then(({ error }) => { if (error) throw error; }), `Promijenio naplatu objekta ${o.name} na ${money(rate, cur)}/h`),
    setBudget: (o, budget, cur) => act(() => supabase.from("object_billing").upsert({ object_id: o.id, monthly_budget: budget })
      .then(({ error }) => { if (error) throw error; }), `Postavio mjesečni proračun objekta ${o.name} na ${money(budget, cur)}`),
    setPositionRate: (o, position, patch) => act(() => supabase.from("position_billing")
      .upsert({ object_id: o.id, position, mode: patch.mode === "markup" ? "markup" : "fixed",
        rate: patch.mode === "markup" ? 0 : (patch.value || 0), markup: patch.mode === "markup" ? (patch.value || 0) : 0,
        currency: patch.currency || "EUR", created_by: session.user.id }, { onConflict: "object_id,position" })
      .then(({ error }) => { if (error) throw error; }),
      `Postavio naplatu za poziciju "${position}" (${o.name}): ${patch.mode === "markup" ? "satnica radnika + " + money(patch.value, patch.currency) : money(patch.value, patch.currency)}/h`),
    delPositionRate: (o, position) => act(() => supabase.from("position_billing").delete().eq("object_id", o.id).eq("position", position)
      .then(({ error }) => { if (error) throw error; }), `Obrisao naplatu za poziciju "${position}" (${o.name})`),
    addLog: async (l, wName, objName) => {
      try { guardPaid(l.workerId, l.date); } catch (e) { setErr(tr("Greška: ") + e.message); return false; }
      if (!l.monthly) {
        const sameDay = data.logs.filter((x) => x.workerId === l.workerId && x.date === l.date && !x.monthly);
        if (sameDay.length && !window.confirm(tr("{0} već ima upisano {1} za {2}.\n\nUpisati još jedan unos za isti dan?", wName, fmtH(round2(sameDay.reduce((s, x) => s + x.hours, 0))), fmtDate(l.date)))) return false;
        if (isDayHoursSuspicious(l.hours) && !window.confirm(tr("{0} za jedan dan je neobično puno.\n\nAko je ovo zbroj za cijeli mjesec, upiši ga u načinu \"Mjesečno\". Svejedno spremiti kao jedan dan?", fmtH(l.hours)))) return false;
      }
      const row = { worker_id: l.workerId, object_id: l.objectId || null, work_date: l.date,
        from_t: l.from, to_t: l.to, hours: l.hours, monthly: !!l.monthly, note: l.note || "", created_by: session.user.id };
      // bez interneta (ili net baš otpao): spremi lokalno i pošalji čim se veza vrati
      if (typeof navigator !== "undefined" && navigator.onLine === false) {
        queuePush("work_log", row);
        setErr("");
        return true;
      }
      setBusy(true);
      try {
        await ins("work_logs", row);
      } catch (e) {
        if (isNetworkError(e)) { queuePush("work_log", row); setErr(""); setBusy(false); return true; }
        setErr(tr("Greška: ") + (e.message || e)); setBusy(false); return false;
      }
      try { await supabase.from("audit_log").insert({ user_id: session.user.id, user_name: profile?.name || "",
        action: `Upisao sate: ${wName} ${fmtH(l.hours)}${objName ? " (" + objName + ")" : ""} (${fmtDate(l.date)})` }); } catch {}
      await reload();
      setErr("");
      setBusy(false);
      return true;
    },
    addLogsBulk: (entries, mo) => act(async () => {
      entries.forEach((e) => guardPaid(e.workerId, mo));
      const { error } = await supabase.from("work_logs").insert(entries.map((e) => ({
        worker_id: e.workerId, object_id: e.objectId || null, work_date: mo + "-01",
        from_t: "", to_t: "", hours: e.hours, monthly: true, created_by: session.user.id,
      })));
      if (error) throw error;
    }, `Upisao mjesečne sate za ${entries.length} radnika (${mo})`),
    // postavlja (ne dodaje) mjesečni zbroj sati po radniku za objekt — jedan, svi ili odabrani odjednom.
    // ako radnik već ima mjesečni unos za taj mjesec, mijenja ga; ako je uneseno 0, briše ga.
    setMonthlyHoursBulk: (entries, objectId, mo, objName) => act(async () => {
      for (const e of entries) {
        guardPaid(e.workerId, mo);
        const existing = data.logs.filter((l) => l.workerId === e.workerId && l.objectId === objectId && l.monthly && monthKey(l.date) === mo);
        if (e.hours > 0) {
          if (existing.length) {
            const { error } = await supabase.from("work_logs").update({ hours: e.hours }).eq("id", existing[0].id);
            if (error) throw error;
            for (const dup of existing.slice(1)) {
              const { error: derr } = await supabase.from("work_logs").update({ deleted_at: new Date().toISOString() }).eq("id", dup.id);
              if (derr) throw derr;
            }
          } else {
            await ins("work_logs", {
              worker_id: e.workerId, object_id: objectId, work_date: mo + "-01",
              from_t: "", to_t: "", hours: e.hours, monthly: true, created_by: session.user.id,
            });
          }
        } else if (existing.length) {
          for (const ex of existing) {
            const { error } = await supabase.from("work_logs").update({ deleted_at: new Date().toISOString() }).eq("id", ex.id);
            if (error) throw error;
          }
        }
      }
    }, `Promijenio mjesečne sate za ${entries.length} radnika (${objName}, ${mo})`),
    delLog: (l, wName) => act(() => { guardPaid(l.workerId, l.date); return softDel("work_logs", l.id); },
      `Obrisao sate: ${wName} ${fmtH(l.hours)} (${fmtDate(l.date)})`),
    updLog: (l, patch, wName) => act(() => {
      guardPaid(l.workerId, l.date);
      guardPaid(l.workerId, patch.date);
      return upd("work_logs", l.id, { work_date: patch.date, from_t: patch.from, to_t: patch.to, hours: patch.hours });
    }, `Ispravio sate: ${wName} ${fmtH(patch.hours)} (${fmtDate(l.date)} → ${fmtDate(patch.date)})`),
    copyDay: (object, fromDate, toDate) => act(async () => {
      const src = data.logs.filter((l) => l.objectId === object.id && l.date === fromDate && !l.monthly);
      if (!src.length) throw new Error(tr("Za {0} nema upisa na ovom objektu.", fmtDate(fromDate)));
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
    addPaymentsBulk: (entries, meta) => act(async () => {
      entries.forEach((e) => guardPaid(e.workerId, meta.date));
      const { error } = await supabase.from("payments").insert(entries.map((e) => ({
        worker_id: e.workerId, object_id: null, pay_date: meta.date, type: meta.type, amount: e.amount,
        note: meta.note || "", deduct: meta.deduct, currency: meta.currency || "EUR", created_by: session.user.id,
      })));
      if (error) throw error;
    }, `Upisao ${TYPE_LABEL[meta.type]?.toLowerCase() || meta.type} za ${entries.length} radnika (${money(round2(entries.reduce((s, e) => s + e.amount, 0)), meta.currency)})`),
    markPaid: (w, mo, amount, amountKc) => act(() => ins("payouts", { worker_id: w.id, month: mo, amount, amount_czk: amountKc || 0, approved: false, created_by: session.user.id }),
      `Predložio isplatu: ${w.name} za ${mo} (${[amount ? eur(amount) : "", amountKc ? czk(amountKc) : ""].filter(Boolean).join(" + ") || eur(0)}) — čeka odobrenje`),
    approvePayout: (payout, wName) => act(() => upd("payouts", payout.id, { approved: true, approved_by: session.user.id, approved_at: new Date().toISOString() }),
      `Odobrio isplatu: ${wName} za ${payout.month}`),
    unmarkPaid: (payout, wName) => act(() => softDel("payouts", payout.id), `Otključao/odbacio isplatu: ${wName} za ${payout.month}`),
    refreshPayout: (payout, amount, amountKc, wName) => act(() => upd("payouts", payout.id, { amount, amount_czk: amountKc || 0 }),
      `Osvježio predloženi iznos isplate: ${wName} za ${payout.month} (sati/isplate su se promijenili nakon prijedloga)`),
    addInvoicePayment: (o, mo, amount) => act(() => ins("invoice_payments", { object_id: o.id, month: mo, amount }),
      `Upisao uplatu od ${o.name}: ${eur(amount)} (${mo})`),
    delInvoicePayment: (p, oName) => act(() => softDel("invoice_payments", p.id), `Obrisao uplatu od ${oName}: ${eur(p.amount)}`),
    importObjectInvoices: (invs) => act(async () => {
      const { error } = await supabase.from("object_invoices").insert(invs.map((inv) => ({
        object_id: inv.objectId, number: inv.number, period: inv.period, issue_date: inv.issueDate, duzp: inv.duzp || null,
        hours: inv.hours || 0, rate: inv.rate || 0, amount: inv.amount, currency: inv.currency || "EUR", note: inv.note || "",
        vat_amount: inv.vatAmount || 0, vat_currency: inv.vatCurrency || "EUR", file_path: inv.filePath || "",
        created_by: session.user.id,
      })));
      if (error) throw error;
    }, `Uvezao ${invs.length} ${invs.length === 1 ? "fakturu" : "fakture"} iz PDF-a`),
    addObjectInvoice: (inv, oName) => act(() => ins("object_invoices", {
      object_id: inv.objectId, number: inv.number, period: inv.period, issue_date: inv.issueDate, duzp: inv.duzp || null,
      hours: inv.hours, rate: inv.rate, amount: inv.amount, currency: inv.currency, note: inv.note || "",
      vat_amount: inv.vatAmount || 0, vat_currency: inv.vatCurrency || inv.currency || "EUR", file_path: inv.filePath || "",
      created_by: session.user.id,
    }), `Generirao fakturu ${inv.number} za ${oName}: ${money(inv.amount, inv.currency)}${inv.vatAmount ? ` (PDV ${money(inv.vatAmount, inv.vatCurrency)})` : ""}`),
    delObjectInvoice: (inv, oName) => act(() => softDel("object_invoices", inv.id), `Obrisao fakturu ${inv.number} (${oName})`),
    setInvoiceFile: (inv, filePath, oName) => act(() => upd("object_invoices", inv.id, { file_path: filePath }),
      `Dodao PDF uz fakturu ${inv.number} (${oName})`),
    delInvoiceFile: (inv, oName) => act(async () => {
      if (inv.filePath) await supabase.storage.from("invoices").remove([inv.filePath]);
      const { error } = await supabase.from("object_invoices").update({ file_path: "" }).eq("id", inv.id);
      if (error) throw error;
    }, `Uklonio PDF uz fakturu ${inv.number} (${oName})`),
    addPayrollNote: (note, wName) => act(() => ins("payroll_notes", {
      worker_id: note.workerId, month: note.month, amount: note.amount, currency: note.currency,
      source_czk: note.sourceCzk, note: note.note || "", created_by: session.user.id,
    }), `Dodao mzdy napomenu: ${wName} ${money(note.amount, note.currency)} (${note.month})`),
    delPayrollNote: (note, wName) => act(() => softDel("payroll_notes", note.id), `Obrisao mzdy napomenu: ${wName} (${note.month})`),
    saveSettings: (s) => {
      const old = data.settings || {};
      const labels = { company_name: "naziv", address: "adresa", oib: "OIB", iban: "IBAN", czk_rate: "tečaj", czk_rate_auto: "automatski tečaj ČNB", weekend_pct: "dodatak vikend", holiday_pct: "dodatak praznik", expiry_warn_days: "dani upozorenja", payday_day: "dan isplate", payday_reminder_days: "dani podsjetnika" };
      const diffs = Object.keys(labels).filter((k) => String(old[k] ?? "") !== String(s[k] ?? "")).map((k) => `${labels[k]} "${old[k] ?? "—"}"→"${s[k] ?? "—"}"`);
      return act(() => supabase.from("settings").upsert({ id: 1, ...s }).then(({ error }) => { if (error) throw error; }),
        `Uredio podatke firme${diffs.length ? " (" + diffs.join(", ") + ")" : ""}`);
    },
    restore: (t) => act(() => upd(t.table, t.row.id, { deleted_at: null }), `Vratio iz koša: ${t.label}`),
    setRole: (p, role) => act(() => upd("profiles", p.id, { role }), `Promijenio ulogu: ${p.name} → ${role === "admin" ? "admin" : "zaposlenik"}`),
    setCommission: (p, rate, cur) => act(() => upd("profiles", p.id, { hourly_commission: rate, commission_currency: cur || "EUR" }),
      `Postavio zadanu proviziju po satu za ${p.name}: ${money(rate, cur)}/h`),
    setCommissionRate: (o, p, rate, cur) => act(() => supabase.from("commission_rates")
      .upsert({ object_id: o.id, user_id: p.id, rate, currency: cur || "EUR", created_by: session.user.id }, { onConflict: "object_id,user_id" })
      .then(({ error }) => { if (error) throw error; }),
      `Postavio proviziju za ${p.name} na objektu ${o.name}: ${money(rate, cur)}/h`),
    delCommissionRate: (o, p) => act(() => supabase.from("commission_rates").delete().eq("object_id", o.id).eq("user_id", p.id)
      .then(({ error }) => { if (error) throw error; }),
      `Obrisao posebnu proviziju za ${p.name} na objektu ${o.name} (vraća se na zadanu)`),
    addCommissionPayout: (p, amount, cur, date, note) => act(() => ins("commission_payouts",
      { user_id: p.id, amount, currency: cur || "EUR", pay_date: date || todayISO(), note: note || "", created_by: session.user.id }),
      `Isplatio proviziju ${p.name}: ${money(amount, cur)}${note ? " (" + note + ")" : ""}`),
    delCommissionPayout: (cp, pName) => act(() => softDel("commission_payouts", cp.id),
      `Obrisao isplatu provizije za ${pName}: ${money(cp.amount, cp.currency)}`),
    addBrigadeCost: (object, b, payerName) => act(() => ins("brigade_costs",
      { object_id: object.id, work_date: b.date, hours: b.hours, amount: b.amount, currency: b.cur || "CZK", note: b.note || "",
        bill_rate: b.billRate || 0, bill_currency: b.billCur || "EUR", people: b.people || 1, names: b.names || "",
        paid_by: b.paidBy || null, created_by: session.user.id }),
      `Upisao brigadu na ${object.name}: ${b.people > 1 ? b.people + " os. · " : ""}${fmtH(b.hours)} za ${money(b.amount, b.cur)} (${fmtDate(b.date)})${b.paidBy ? ` — platio ${payerName} iz gotovine` : ""}${b.billRate > 0 ? ` — posebna cijena naplate ${money(b.billRate, b.billCur)}/h` : ""}`),
    delBrigadeCost: (b, objName) => act(() => softDel("brigade_costs", b.id),
      `Obrisao brigadu na ${objName}: ${fmtH(b.hours)} / ${money(b.amount, b.currency)} (${fmtDate(b.date)})`),
    addCashAdvance: (p, kind, amount, cur, date, note) => act(() => ins("cash_advances",
      { profile_id: p.id, kind, amount, currency: cur || "CZK", pay_date: date || todayISO(), note: note || "", created_by: session.user.id }),
      `${kind === "returned" ? "Koordinator " + p.name + " vratio gotovinu" : "Dao gotovinu koordinatoru " + p.name}: ${money(amount, cur)} (${fmtDate(date)})${note ? " — " + note : ""}`),
    delCashAdvance: (c, pName) => act(() => softDel("cash_advances", c.id),
      `Obrisao zapis gotovine (${pName}): ${money(c.amount, c.currency)} (${fmtDate(c.date)})`),
    toggleMember: (object, p, on) => act(() =>
      (on ? supabase.from("object_members").insert({ object_id: object.id, user_id: p.id })
          : supabase.from("object_members").delete().eq("object_id", object.id).eq("user_id", p.id)
      ).then(({ error }) => { if (error) throw error; }),
      on ? `Dodijelio objekt ${object.name} zaposleniku ${p.name}` : `Maknuo objekt ${object.name} zaposleniku ${p.name}`),
    reload: () => reload(),
    setCloseStep: (month, objectId, step, on, label) => act(() => (on
      ? supabase.from("month_close").upsert({ month, object_id: objectId || null, step, done_by: session.user.id, done_at: new Date().toISOString() }, { onConflict: "month,object_id,step" })
      : (objectId ? supabase.from("month_close").delete().eq("month", month).eq("object_id", objectId).eq("step", step)
                  : supabase.from("month_close").delete().eq("month", month).is("object_id", null).eq("step", step))
    ).then(({ error }) => { if (error) throw error; }),
      `Zatvaranje ${month}: ${label || ""} — ${step === "done" ? "objekt" : step} ${on ? "✓" : "otvoreno ponovno"}`),
    addPurchaseInvoice: (f) => act(() => ins("purchase_invoices", {
      supplier: f.supplier, number: f.number, duzp: f.duzp, base_amount: f.base, vat_amount: f.vat, currency: f.currency || "CZK", note: f.note || "",
      created_by: session.user.id,
    }), `Upisao primljenu fakturu: ${f.supplier} ${f.number} (DPH ${money(f.vat, f.currency)}, DUZP ${fmtDate(f.duzp)})`),
    delPurchaseInvoice: (pi) => act(() => softDel("purchase_invoices", pi.id),
      `Obrisao primljenu fakturu: ${pi.supplier} ${pi.number} (DPH ${money(pi.vat, pi.currency)})`),
    createEmployee: async (f) => {
      setBusy(true);
      try {
        const { data: res, error } = await supabase.functions.invoke("create-employee", { body: f });
        if (error) {
          let msg = error.message;
          try { const j = await error.context.json(); if (j?.error) msg = j.error; } catch {}
          throw new Error(msg);
        }
        if (res?.error) throw new Error(res.error);
        await reload();
        setErr(res?.warning ? "Upozorenje: " + res.warning : "");
        return true;
      } catch (e) { setErr(tr("Greška: ") + (e.message || e)); return false; }
      finally { setBusy(false); }
    },
    logDoc: (text) => supabase.from("audit_log").insert({ user_id: session.user.id, user_name: profile?.name || "", action: text }),
    addReminder: (text, dueDate) => act(() => ins("reminders", { text, due_date: dueDate || null, created_by: session.user.id }),
      `Dodao podsjetnik: ${text}${dueDate ? " (" + fmtDate(dueDate) + ")" : ""}`),
    toggleReminder: (r, done) => act(() => upd("reminders", r.id, { done }),
      done ? `Označio podsjetnik kao gotov: ${r.text}` : `Vratio podsjetnik: ${r.text}`),
    delReminder: (r) => act(() => softDel("reminders", r.id), `Obrisao podsjetnik: ${r.text}`),
  };

  if (session === undefined) return <Center>{tr("Učitavam…")}</Center>;
  if (!session) return <Login lang={lang} setLang={setLang} t={t} />;
  if (!data) return <Center>{err || tr("Učitavam podatke…")}</Center>;

  const worker = data.workers.find((w) => w.id === openWorker);
  const object = data.objects.find((o) => o.id === openObject);
  const openedProfile = data.profiles.find((p) => p.id === openProfile);
  const detailOpen = worker || object || openedProfile;

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

      <Header profile={profile} admin={admin} onLogout={() => supabase.auth.signOut()} lang={lang} setLang={setLang} t={t} data={data}
        title={adminPanel ? (ADMIN_MENU.find((m) => m[0] === adminPanel) || [])[2] : worker ? worker.name : object ? object.name : openedProfile ? openedProfile.name : tab === "kandidati" ? tr("Kandidati") : t(TAB_TITLE_KEY[tab] || "tabDanas")}
        onAdminPanel={(p) => { setAdminPanel(p); setOpenWorker(null); setOpenObject(null); setOpenProfile(null); window.scrollTo({ top: 0, behavior: "smooth" }); }} />

      <div style={{ maxWidth: 560, margin: "0 auto", padding: "0 14px" }}>
        {err && (
          <div style={{ background: S.redSoft, color: S.red, borderRadius: 10, padding: "9px 12px", marginBottom: 12, fontSize: 13.5, fontWeight: 600 }}>
            {err} <button onClick={() => setErr("")} style={{ background: "none", border: "none", color: S.red, fontWeight: 800, cursor: "pointer", float: "right" }}>✕</button>
          </div>
        )}
        {busy && <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 8 }}>{tr("Spremam…")}</div>}
        {offlineQueue.length > 0 && (
          <div style={{ background: S.amberSoft, color: S.amber, borderRadius: 10, padding: "9px 12px", marginBottom: 12, fontSize: 13, fontWeight: 600,
            display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
            <span>📴 {offlineQueue.length} {offlineQueue.length === 1 ? tr("upis čeka") : tr("upisa čeka")} {tr("slanje (nema interneta)")}</span>
            <button onClick={flushOfflineQueue} style={{ background: "none", border: `1px solid ${S.amber}`, color: S.amber, borderRadius: 8, padding: "4px 9px", fontWeight: 700, cursor: "pointer", fontSize: 12.5 }}>
              {tr("Pokušaj sad")}
            </button>
          </div>
        )}

        {admin && !detailOpen && adminPanel && <AdminPanels data={data} api={api} panel={adminPanel} setPanel={setAdminPanel} onOpenWorker={(id) => { setAdminPanel(""); setOpenWorker(id); }} onOpenProfile={(id) => { setAdminPanel(""); setOpenProfile(id); }} />}

        {worker ? (
          <WorkerDetail worker={worker} data={data} api={api} onBack={() => setOpenWorker(null)} />
        ) : object ? (
          <ObjectDetail object={object} data={data} api={api} onBack={() => setOpenObject(null)} onOpenWorker={setOpenWorker} />
        ) : openedProfile ? (
          <ProfileDetail profile={openedProfile} data={data} api={api} onBack={() => setOpenProfile(null)} />
        ) : adminPanel ? null : (
          <>
            {tab === "danas" && <TodayTab data={data} api={api} admin={admin} onOpenWorker={setOpenWorker} onOpenObject={setOpenObject}
              openAdminPanel={(p) => { setAdminPanel(p); window.scrollTo({ top: 0, behavior: "smooth" }); }} goTab={setTab} />}
            {tab === "kandidati" && api.recruiter && <CandidatesPanel data={data} api={api} />}
            {tab === "radnici" && <WorkersTab data={data} api={api} onOpen={setOpenWorker} onOpenObject={setOpenObject} />}
            {tab === "objekti" && <ObjectsTab data={data} api={api} onOpenObject={setOpenObject} />}
            {tab === "imenik" && <DirectoryTab data={data} api={api} onOpen={setOpenWorker} />}
            {tab === "sati" && <HoursTab data={data} api={api} onOpenWorker={setOpenWorker} />}
            {tab === "isplate" && <PaymentsTab data={data} api={api} onOpenWorker={setOpenWorker} />}
            {tab === "obracun" && <ReportTab data={data} api={api} admin={admin} onOpenWorker={setOpenWorker} />}
          </>
        )}
      </div>

      {!detailOpen && <TabBar tab={adminPanel ? "" : tab} setTab={(x) => { setAdminPanel(""); setTab(x); window.scrollTo({ top: 0 }); }} t={t} recruiter={api.recruiter} newCandidates={data?.newCandidates || 0} />}
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
  const loginGoogle = async () => {
    setErr("");
    const { error } = await supabase.auth.signInWithOAuth({ provider: "google", options: { redirectTo: window.location.origin } });
    if (error) setErr(tr("Google prijava nije uključena — treba se postaviti u Supabase Dashboardu (Authentication → Providers → Google)."));
  };
  return (
    <div style={{ minHeight: "100vh", background: S.bg, fontFamily: "'Segoe UI', system-ui, sans-serif", color: S.ink }}>
      <style>{`* { box-sizing: border-box; } input { width: 100%; padding: 12px; border: 1px solid ${S.line}; border-radius: 10px; outline: none; font-size: 15px; } input:focus { border-color: ${S.green}; }`}</style>
      <div style={{ background: S.ink, color: "#fff", padding: "26px 16px" }}>
        <div style={{ maxWidth: 420, margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <BrandMark size={40} />
              <div>
                <div style={{ fontSize: 20, fontWeight: 800 }}>{BRAND}</div>
                <div style={{ fontSize: 12.5, color: "#9DB3A8" }}>{t("appTitle")} · {t("loginTitle")}</div>
              </div>
            </div>
          </div>
          <LangSwitch lang={lang} setLang={setLang} dark />
        </div>
      </div>
      <div style={{ maxWidth: 420, margin: "26px auto 0", padding: "0 16px" }}>
        <div style={{ background: "#fff", border: `1px solid ${S.line}`, borderRadius: 14, padding: 18 }}>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: S.sub, marginBottom: 5 }}>{t("email")}</div>
          <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder={tr("ime@firma.hr")} autoComplete="username" />
          <div style={{ fontSize: 12.5, fontWeight: 600, color: S.sub, margin: "12px 0 5px" }}>{t("password")}</div>
          <input type="password" value={pass} onChange={(e) => setPass(e.target.value)} autoComplete="current-password" onKeyDown={(e) => e.key === "Enter" && login()} />
          {err && <div style={{ color: S.red, fontSize: 13, marginTop: 10 }}>{err}</div>}
          <button onClick={login} disabled={busy} style={{ width: "100%", background: S.green, color: "#fff", border: "none", borderRadius: 10, padding: 13, fontWeight: 700, cursor: "pointer", marginTop: 14, fontSize: 15, opacity: busy ? 0.7 : 1 }}>
            {busy ? t("loginBusy") : t("loginBtn")}
          </button>
          <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "14px 0" }}>
            <div style={{ flex: 1, height: 1, background: S.line }} />
            <span style={{ fontSize: 12, color: S.sub }}>{tr("ili")}</span>
            <div style={{ flex: 1, height: 1, background: S.line }} />
          </div>
          <button onClick={loginGoogle} style={{ width: "100%", background: "#fff", color: S.ink, border: `1px solid ${S.line}`, borderRadius: 10, padding: 12, fontWeight: 700, cursor: "pointer", fontSize: 14.5 }}>
            <span style={{ display: "inline-flex", alignItems: "center", gap: 8, justifyContent: "center" }}>
              <svg width="18" height="18" viewBox="0 0 48 48" aria-hidden="true"><path fill="#FFC107" d="M43.6 20.5H42V20H24v8h11.3C33.7 32.7 29.2 36 24 36c-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.8 1.2 7.9 3.1l5.7-5.7C34 6.1 29.3 4 24 4 12.9 4 4 12.9 4 24s8.9 20 20 20 20-8.9 20-20c0-1.3-.1-2.4-.4-3.5z"/><path fill="#FF3D00" d="M6.3 14.7l6.6 4.8C14.7 15.1 19 12 24 12c3.1 0 5.8 1.2 7.9 3.1l5.7-5.7C34 6.1 29.3 4 24 4 16.3 4 9.7 8.3 6.3 14.7z"/><path fill="#4CAF50" d="M24 44c5.2 0 9.9-2 13.4-5.2l-6.2-5.2C29.2 35.1 26.7 36 24 36c-5.2 0-9.6-3.3-11.3-8l-6.5 5C9.5 39.6 16.2 44 24 44z"/><path fill="#1976D2" d="M43.6 20.5H42V20H24v8h11.3c-.8 2.2-2.2 4.2-4.1 5.6l6.2 5.2C37 39.2 44 34 44 24c0-1.3-.1-2.4-.4-3.5z"/></svg>
              {tr("Prijavi se Google računom")}
            </span>
          </button>
          <div style={{ fontSize: 12.5, color: S.sub, marginTop: 14, lineHeight: 1.5 }}>
            {t("noAccount")}
          </div>
        </div>
      </div>
    </div>
  );
}

function Header({ profile, admin, onLogout, lang, setLang, t, title, data, onAdminPanel }) {
  const [menu, setMenu] = useState(""); // "" | "admin" | "user"
  const pending = admin ? adminBadge(data, "approvals") : 0;
  const close = () => setMenu("");
  const menuBox = { position: "absolute", right: 0, top: "calc(100% + 8px)", background: "#fff", color: S.ink, borderRadius: 14,
    boxShadow: "0 10px 30px rgba(0,0,0,.18)", border: `1px solid ${S.line}`, minWidth: 250, padding: 6, zIndex: 50 };
  const item = { display: "flex", alignItems: "center", gap: 10, width: "100%", background: "none", border: "none", textAlign: "left",
    padding: "10px 10px", borderRadius: 10, cursor: "pointer", fontSize: 14, color: S.ink };
  return (
    <div style={{ background: S.ink, color: "#fff", padding: "12px 16px", marginBottom: 16, position: "sticky", top: 0, zIndex: 30 }}>
      {menu && <div onClick={close} style={{ position: "fixed", inset: 0, zIndex: 40 }} />}
      <div style={{ maxWidth: 560, margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
          <BrandMark />
          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: 11, letterSpacing: 1.5, textTransform: "uppercase", color: "#9DB3A8", fontWeight: 700 }}>{BRAND}</div>
            <div style={{ fontSize: 19, fontWeight: 800, lineHeight: 1.15, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{title}</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, position: "relative" }}>
          {admin && (
            <button onClick={() => setMenu(menu === "admin" ? "" : "admin")} title={tr("Upravljanje")} style={{ position: "relative", background: menu === "admin" ? "rgba(255,255,255,.22)" : "rgba(255,255,255,.1)",
              border: "1px solid rgba(255,255,255,.2)", borderRadius: 10, color: "#fff", padding: "7px 11px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
              {tr("⚙︎ Upravljanje")}
              {pending > 0 && <span style={{ position: "absolute", top: -6, right: -6, background: S.amber, color: "#fff", borderRadius: 999, fontSize: 11, padding: "1px 6px", fontWeight: 800 }}>{pending}</span>}
            </button>
          )}
          <button onClick={() => setMenu(menu === "user" ? "" : "user")} title={profile?.name || ""} style={{ width: 36, height: 36, borderRadius: 999, background: "rgba(255,255,255,.14)",
            border: "1px solid rgba(255,255,255,.25)", color: "#fff", fontWeight: 800, fontSize: 13, cursor: "pointer" }}>
            {(profile?.name || "?").split(" ").map((x) => x[0]).join("").slice(0, 2).toUpperCase()}
          </button>
          {menu === "admin" && (
            <div style={{ ...menuBox, zIndex: 50 }}>
              {ADMIN_MENU.map(([id, icon, label]) => {
                const n = adminBadge(data, id);
                return (
                  <button key={id} onClick={() => { onAdminPanel(id); close(); }} style={item}
                    onMouseEnter={(e) => (e.currentTarget.style.background = S.bg)} onMouseLeave={(e) => (e.currentTarget.style.background = "none")}>
                    <span style={{ width: 22, textAlign: "center" }}>{icon}</span>
                    <span style={{ flex: 1 }}>{label}</span>
                    {n > 0 && <span style={{ background: id === "approvals" || id === "candidates" ? S.amberSoft : "#EEF0ED", color: id === "approvals" || id === "candidates" ? S.amber : S.sub, borderRadius: 999, fontSize: 12, fontWeight: 700, padding: "1px 8px" }}>{n}</span>}
                  </button>
                );
              })}
            </div>
          )}
          {menu === "user" && (
            <div style={{ ...menuBox, zIndex: 50 }}>
              <div style={{ padding: "8px 10px 10px", borderBottom: `1px solid ${S.line}`, marginBottom: 6 }}>
                <div style={{ fontWeight: 700 }}>{profile?.name || "—"}</div>
                <div style={{ fontSize: 12, color: S.sub }}>{admin ? tr("Administrator") : tr("Zaposlenik")}</div>
              </div>
              <div style={{ padding: "4px 10px 8px" }}>
                <div style={{ fontSize: 12, color: S.sub, marginBottom: 6 }}>{tr("Jezik")}</div>
                <LangSwitch lang={lang} setLang={setLang} labeled />
              </div>
              <button onClick={onLogout} style={{ ...item, color: S.red, fontWeight: 700 }}>↪ {t("logout")}</button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
const TAB_TITLE_KEY = { danas: "tabDanas", radnici: "tabRadnici", objekti: "tabObjekti", imenik: "tabImenik", sati: "tabSati", isplate: "tabIsplate", obracun: "tabObracun" };
function TabBar({ tab, setTab, t, recruiter, newCandidates }) {
  const tabs = [["danas",t("tabDanas"),"📌"],["radnici",t("tabRadnici"),"👷"],["objekti",t("tabObjekti"),"🏨"],["imenik",t("tabImenik"),"📇"],["sati",t("tabSati"),"🕒"],["isplate",t("tabIsplate"),"💶"],["obracun",t("tabObracun"),"🧾"]];
  // regruteri i admin imaju i karticu Kandidati (s brojem novih prijava)
  if (recruiter) tabs.splice(1, 0, ["kandidati", tr("Kandidati") + (newCandidates ? ` (${newCandidates})` : ""), "🧑‍🍳"]);
  return (
    <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: "#fff", borderTop: `1px solid ${S.line}`,
      display: "flex", justifyContent: "space-around", padding: "6px 4px calc(6px + env(safe-area-inset-bottom))", zIndex: 20 }}>
      {tabs.map(([id, label, icon]) => (
        <button key={id} onClick={() => setTab(id)} style={{ background: "none", border: "none", padding: "6px 3px", borderRadius: 10, minWidth: 0, flex: 1,
          color: tab === id ? S.green : S.sub, fontWeight: tab === id ? 700 : 500, fontSize: 11,
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
function SignaturePad({ title, onSave, onClose }) {
  const canvasRef = useRef(null);
  const drawing = useRef(false);
  const last = useRef({ x: 0, y: 0 });
  const [empty, setEmpty] = useState(true);
  useEffect(() => {
    const ctx = canvasRef.current.getContext("2d");
    ctx.lineWidth = 2.5; ctx.lineCap = "round"; ctx.lineJoin = "round"; ctx.strokeStyle = S.ink;
  }, []);
  const pos = (e) => {
    const c = canvasRef.current;
    const rect = c.getBoundingClientRect();
    const t = e.touches ? e.touches[0] : e;
    return { x: (t.clientX - rect.left) * (c.width / rect.width), y: (t.clientY - rect.top) * (c.height / rect.height) };
  };
  const start = (e) => { e.preventDefault(); drawing.current = true; last.current = pos(e); };
  const move = (e) => {
    if (!drawing.current) return;
    e.preventDefault();
    const ctx = canvasRef.current.getContext("2d");
    const p = pos(e);
    ctx.beginPath(); ctx.moveTo(last.current.x, last.current.y); ctx.lineTo(p.x, p.y); ctx.stroke();
    last.current = p;
    setEmpty(false);
  };
  const end = () => { drawing.current = false; };
  const clear = () => { const c = canvasRef.current; c.getContext("2d").clearRect(0, 0, c.width, c.height); setEmpty(true); };
  const save = () => { if (empty) return; onSave(canvasRef.current.toDataURL("image/png")); };
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(10,14,12,.92)", zIndex: 60,
      display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", borderRadius: 14, padding: 16, maxWidth: 420, width: "100%" }}>
        <div style={{ fontWeight: 700, marginBottom: 8 }}>✍️ {title || tr("Potpis")}</div>
        <canvas ref={canvasRef} width={380} height={180}
          onMouseDown={start} onMouseMove={move} onMouseUp={end} onMouseLeave={end}
          onTouchStart={start} onTouchMove={move} onTouchEnd={end}
          style={{ width: "100%", height: 180, border: `1px solid ${S.line}`, borderRadius: 10, touchAction: "none", background: "#FAFBF9", display: "block" }} />
        <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
          <Btn small kind="ghost" onClick={clear}>{tr("Obriši")}</Btn>
          <Btn small onClick={save} disabled={empty} style={{ flex: 1 }}>{tr("✓ Spremi potpis")}</Btn>
          <Btn small kind="ghost" onClick={onClose}>{tr("Zatvori")}</Btn>
        </div>
      </div>
    </div>
  );
}
function Tag({ children, color, bg }) {
  return <span style={{ background: bg, color, borderRadius: 999, padding: "3px 10px", fontSize: 12, fontWeight: 700, whiteSpace: "nowrap" }}>{children}</span>;
}
// iznos u dvije valute: veći je onaj koji postoji (samo Kč → Kč gore), drugi ispod manjim slovima
function TwoCur({ e, k, color }) {
  const czkMain = e === 0 && k !== 0;
  return (
    <>
      <div className="num" style={{ fontSize: 18, fontWeight: 800, color }}>{czkMain ? czk(k) : eur(e)}</div>
      {!czkMain && k !== 0 && <div className="num" style={{ fontSize: 13.5, fontWeight: 700, color }}>{czk(k)}</div>}
    </>
  );
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
        <input inputMode="decimal" value={f.hours} onChange={(e) => setF({ ...f, hours: e.target.value })} placeholder={tr("sati")} style={{ flex: "1 1 60px", padding: "8px" }} />
      </div>
      <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
        <Btn small onClick={save}>{tr("Spremi ispravak")}</Btn>
        <Btn small kind="ghost" onClick={onCancel}>{tr("Odustani")}</Btn>
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
      <input autoFocus value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && commit()} placeholder={tr("Naziv objekta")} />
      <Btn small onClick={commit}>{tr("Dodaj")}</Btn>
      <Btn small kind="ghost" onClick={() => { setAdding(false); setName(""); }}>✕</Btn>
    </div>
  );
  return (
    <select value={value} onChange={(e) => (e.target.value === "__new__" ? setAdding(true) : onChange(e.target.value))}>
      <option value="">{tr("— bez objekta —")}</option>
      {["HR", "CZ"].map((c) => {
        const objs = sortedObjects(data.objects.filter((o) => (o.country || "HR") === c && (!o.archived || o.id === value)));
        return objs.length ? (
          <optgroup key={c} label={`${FLAG[c]} ${COUNTRY_NAME[c]}`}>
            {objs.map((o) => <option key={o.id} value={o.id}>{o.name}</option>)}
          </optgroup>
        ) : null;
      })}
      <option value="__new__">{tr("＋ Dodaj novi objekt…")}</option>
    </select>
  );
}

function exportAllData(data) {
  const objName = (id) => data.objects.find((o) => o.id === id)?.name || "";
  const wName = (id) => data.workers.find((w) => w.id === id)?.name || "";
  const cur = (c) => (c === "CZK" ? "Kč" : "€");
  const byDate = (a, b) => (a.date || "").localeCompare(b.date || "");
  const statusColor = (v) => (v === "DA" ? XC.amber : null);
  return downloadStyledXlsx({
    fileName: `Backup_${todayISO()}.xlsx`,
    sheets: [
      { name: "Radnici", title: "Radnici", subtitle: "Kopija svih podataka · Evidencija rada", landscape: true, freezeCols: 1,
        columns: [
          { key: "name", header: "Prezime i ime", width: 26, bold: true }, { key: "phone", header: "Telefon", width: 16 },
          { key: "obj", header: "Objekt", width: 20 }, { key: "pos", header: "Pozicija", width: 18 },
          { key: "rate", header: "Satnica", width: 10, type: "money" }, { key: "cur", header: "Valuta", width: 8, type: "center" },
          { key: "arch", header: "Bivši radnik", width: 11, type: "center", color: statusColor },
          { key: "permit", header: "Istek dozvole", width: 13, type: "date" }, { key: "contract", header: "Istek ugovora", width: 13, type: "date" },
          { key: "note", header: "Napomena", width: 34, wrap: true },
        ],
        rows: sortedWorkers(data.workers).map((w) => ({ name: w.name, phone: w.phone || "", obj: objName(w.objectId), pos: w.position || "",
          rate: w.rate || "", cur: cur(w.rateCurrency), arch: w.archived ? "DA" : "", permit: w.permitExpiry, contract: w.contractExpiry, note: w.note || "" })) },
      { name: "Objekti", title: "Objekti", subtitle: "Kopija svih podataka · Evidencija rada",
        columns: [
          { key: "name", header: "Naziv", width: 26, bold: true }, { key: "country", header: "Država", width: 9, type: "center" },
          { key: "rate", header: "Naplata po satu", width: 15, type: "money" }, { key: "cur", header: "Valuta", width: 9, type: "center" },
          { key: "arch", header: "Arhiviran", width: 10, type: "center", color: statusColor },
        ],
        rows: sortedObjects(data.objects).map((o) => ({ name: o.name, country: o.country, rate: o.billRate || "", cur: cur(o.billCur), arch: o.archived ? "DA" : "" })) },
      { name: "Sati", title: "Upisani sati", subtitle: "Kopija svih podataka · Evidencija rada", landscape: true, total: { keys: ["hours"] },
        columns: [
          { key: "date", header: "Datum", width: 12, type: "date" }, { key: "w", header: "Radnik", width: 26, bold: true },
          { key: "obj", header: "Objekt", width: 20 }, { key: "from", header: "Od", width: 8, type: "center" }, { key: "to", header: "Do", width: 8, type: "center" },
          { key: "hours", header: "Sati", width: 9, type: "hours" }, { key: "monthly", header: "Mjesečni zbroj", width: 11, type: "center" },
          { key: "note", header: "Napomena", width: 34, wrap: true },
        ],
        rows: [...data.logs].sort(byDate).map((l) => ({ date: l.date, w: wName(l.workerId), obj: objName(l.objectId), from: l.from || "", to: l.to || "",
          hours: l.hours, monthly: l.monthly ? "DA" : "", note: l.note || "" })) },
      { name: "Isplate", title: "Isplate i troškovi", subtitle: "Kopija svih podataka · Evidencija rada", landscape: true,
        columns: [
          { key: "date", header: "Datum", width: 12, type: "date" }, { key: "w", header: "Radnik", width: 26, bold: true },
          { key: "obj", header: "Objekt", width: 20 }, { key: "type", header: "Vrsta", width: 14 },
          { key: "amount", header: "Iznos", width: 12, type: "money" }, { key: "cur", header: "Valuta", width: 8, type: "center" },
          { key: "deduct", header: "Odbija se", width: 10, type: "center" }, { key: "note", header: "Napomena", width: 34, wrap: true },
        ],
        rows: [...data.payments].sort(byDate).map((p) => ({ date: p.date, w: p.workerId ? wName(p.workerId) : "", obj: p.objectId ? objName(p.objectId) : "",
          type: TYPE_LABEL[p.type] || p.type, amount: p.amount, cur: cur(p.currency), deduct: p.deduct ? "DA" : "", note: p.note || "" })) },
      { name: "Isplaćeni mjeseci", title: "Isplaćeni mjeseci", subtitle: "Kopija svih podataka · Evidencija rada", total: { keys: ["eur", "czk"] },
        columns: [
          { key: "w", header: "Radnik", width: 26, bold: true }, { key: "month", header: "Mjesec", width: 10, type: "center" },
          { key: "eur", header: "Iznos (€)", width: 13, type: "eur" }, { key: "czk", header: "Iznos (Kč)", width: 13, type: "czk" },
          { key: "date", header: "Datum isplate", width: 13, type: "date" },
          { key: "status", header: "Status", width: 13, type: "center", color: (v) => (v === "Odobreno" ? XC.green : XC.amber) },
        ],
        rows: [...(data.payouts || [])].sort((a, b) => a.month.localeCompare(b.month) || wName(a.workerId).localeCompare(wName(b.workerId), "hr")).map((p) => ({
          w: wName(p.workerId), month: p.month, eur: p.amount || 0, czk: p.amountKc || 0, date: p.paidAt, status: p.approved ? "Odobreno" : "Predloženo" })) },
      { name: "Cijene po poziciji", title: "Cijene po poziciji", subtitle: "Kopija svih podataka · Evidencija rada",
        columns: [
          { key: "obj", header: "Objekt", width: 22, bold: true }, { key: "pos", header: "Pozicija", width: 22 },
          { key: "mode", header: "Način", width: 26 }, { key: "amount", header: "Iznos", width: 12, type: "money" },
          { key: "cur", header: "Valuta", width: 8, type: "center" },
        ],
        rows: (data.positionBilling || []).map((pb) => ({ obj: objName(pb.objectId), pos: pb.position,
          mode: pb.mode === "markup" ? "satnica radnika + dodatak" : "fiksna cijena", amount: pb.mode === "markup" ? pb.markup : pb.rate, cur: cur(pb.currency) })) },
    ],
  }).catch((e) => alert("Izvoz nije uspio: " + (e.message || e)));
}

/* ================================================================== */
/*  ADMIN PANELI                                                       */
/* ================================================================== */
/* ================================================================== */
/*  KANDIDATI — prijave s javne prijavnice /prijava                   */
/*  Tablica: candidates (unos kroz funkciju submit_application)       */
/* ================================================================== */
// Tekstovi oglasa žive u tablici ad_texts (uređuju admin i regruteri); ovo je samo rezerva ako se ne učitaju.
const AD_FALLBACK = [
  { id: "fallback", label: "Sve pozicije", body: "🏨 Робота в готелях Праги\n\nШукаємо покоївок, кухарів, помічників кухаря та мийників посуду.\n📍 Прага\n\n👉 Залиште заявку за 1 хвилину – ми вам зателефонуємо:", is_cs: false },
];
const LL = () => (CUR_LANG === "UK" ? "UK" : "HR");
const digitsPhone = (p) => (p || "").replace(/[^0-9+]/g, "");
const phoneKey = (p) => (p || "").replace(/[^0-9]/g, "").slice(-9); // zadnjih 9 znamenki (bez pozivnog)
const hoursSince = (iso) => (iso ? (Date.now() - new Date(iso).getTime()) / 3600000 : 0);
const slugify = (s) => (s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9а-яіїєґ]+/gi, "-").replace(/^-|-$/g, "").slice(0, 40);
// prva poruka kandidatu, na jeziku na kojem je ispunio prijavnicu
const firstMessage = (c, myName) => {
  const first = (c.full_name || "").trim().split(/\s+/)[0] || "";
  const csLang = c.lang === "CS";
  const pos = (c.positions || []).map((p) => labelOf(POSITIONS, p, csLang ? "CS" : "UK").toLowerCase()).join(", ");
  return csLang
    ? `Dobrý den${first ? `, ${first}` : ""}! Vyplnil(a) jste dotazník na práci v hotelu v Praze (${pos}).${myName ? ` Jmenuji se ${myName}.` : ""} Kdy vám mohu zavolat a domluvit podrobnosti?`
    : `Добрий день${first ? `, ${first}` : ""}! Ви залишили заявку на роботу в готелі в Празі (${pos}).${myName ? ` Мене звати ${myName}.` : ""} Коли вам зручно поговорити, щоб я розповіла деталі?`;
};
// rok za javljanje: novi > 24 h, kontaktiran/razgovor bez pomaka > 3 dana
const staleInfo = (c) => {
  if (c.status === "novi") { const h = hoursSince(c.created_at); return h >= 24 ? { level: "red", text: `${tr("čeka")} ${Math.floor(h / 24) >= 1 ? plural(Math.floor(h / 24), "dan", "dana", "dana") : Math.floor(h) + " h"}` } : null; }
  if (c.status === "kontaktiran" || c.status === "razgovor") {
    const h = hoursSince(c.status_changed_at || c.created_at);
    return h >= 72 ? { level: "amber", text: `${plural(Math.floor(h / 24), "dan", "dana", "dana")} ${tr("bez pomaka")}` } : null;
  }
  return null;
};
const contactLinks = (c, msg) => {
  const ph = digitsPhone(c.phone);
  const intl = ph.replace(/^\+/, "");
  return [
    { icon: "📞", label: tr("Nazovi"), href: `tel:${ph}` },
    { icon: "🟢", label: tr("WhatsApp + poruka"), href: `https://wa.me/${intl}?text=${encodeURIComponent(msg)}` },
    { icon: "💬", label: tr("Viber + poruka"), href: `viber://chat?number=%2B${intl}`, copy: msg },
    { icon: "✈️", label: "Telegram", href: `https://t.me/+${intl}` },
  ];
};
function CandidatesPanel({ data, api }) {
  const [rows, setRows] = useState(null);
  const [loadErr, setLoadErr] = useState("");
  const [filter, setFilter] = useState("aktivni"); // aktivni | svi | <status>
  const [pos, setPos] = useState("");
  const [open, setOpen] = useState(null);
  const [notes, setNotes] = useState({});
  const [copied, setCopied] = useState("");
  const [refWorker, setRefWorker] = useState("");
  const [adId, setAdId] = useState("");
  const [ads, setAds] = useState(null);
  const [adEdit, setAdEdit] = useState(null); // { id|null, label, body, is_cs }
  const [adTag, setAdTag] = useState(() => { try { return localStorage.getItem("evidencija_ad_tag") || ""; } catch { return ""; } });
  const formUrl = `${window.location.origin}/robota`;

  const load = useCallback(async () => {
    const { data: d, error } = await supabase.from("candidates").select("*").order("created_at", { ascending: false }).limit(1000);
    if (error) { setLoadErr(error.message); return; }
    setLoadErr(""); setRows(d || []);
    const { data: a } = await supabase.from("ad_texts").select("*").order("sort").order("created_at");
    setAds(a || []);
  }, []);
  useEffect(() => { load(); }, [load]);

  const objects = sortedObjects(data.objects.filter((o) => !o.archived));
  const activeWorkers = sortedWorkers(data.workers.filter((w) => !w.archived));
  const objName = (id) => data.objects.find((o) => o.id === id)?.name || "";
  const workerByPhone = useMemo(() => {
    const m = {};
    data.workers.forEach((w) => { const k = phoneKey(w.phone); if (k.length === 9) m[k] = w; });
    return m;
  }, [data.workers]);
  const phoneCount = useMemo(() => {
    const m = {};
    (rows || []).forEach((c) => { const k = phoneKey(c.phone); m[k] = (m[k] || 0) + 1; });
    return m;
  }, [rows]);

  const update = async (c, patch, audit) => {
    const ok = await api.run(async () => {
      const { error } = await supabase.from("candidates").update({ ...patch, updated_by: api.uid() }).eq("id", c.id);
      if (error) throw error;
    }, audit);
    if (ok) setRows((prev) => prev.map((r) => (r.id === c.id ? { ...r, ...patch } : r)));
  };
  const setStatus = (c, status) => update(c, { status, status_changed_at: new Date().toISOString() },
    `Kandidat ${c.full_name}: status → ${CAND_STATUS.find((s) => s.id === status)?.label || status}`);
  const remove = async (c) => {
    if (!window.confirm(tr("Trajno obrisati prijavu ovog kandidata?"))) return;
    const ok = await api.run(async () => {
      const { error } = await supabase.from("candidates").delete().eq("id", c.id);
      if (error) throw error;
    }, `Obrisana prijava kandidata: ${c.full_name}`);
    if (ok) setRows((prev) => prev.filter((r) => r.id !== c.id));
  };
  // kandidat počinje raditi → novi radnik (ime, telefon, hotel, pozicija) i veza na prijavu
  const toWorker = async (c) => {
    if (!window.confirm(tr("Dodati kandidata kao radnika u karticu Radnici?"))) return;
    const obj = data.objects.find((o) => o.id === c.object_id);
    let workerId = null;
    const ok = await api.run(async () => {
      const { data: w, error } = await supabase.from("workers").insert({
        name: normalizeName(c.full_name), phone: c.phone, base_rate: 0, rate_currency: obj ? countryCur(obj.country) : "CZK",
        object_id: c.object_id || null, position: labelOf(POSITIONS, (c.positions || [])[0], "HR"), note: "",
        created_by: api.uid(),
      }).select().single();
      if (error) throw error;
      workerId = w.id;
      if (c.object_id) {
        const { error: e2 } = await supabase.from("assignments").insert({ worker_id: w.id, object_id: c.object_id, from_date: todayISO(), created_by: api.uid() });
        if (e2) throw e2;
      }
      const { error: e3 } = await supabase.from("candidates").update({ worker_id: w.id, status: "poceo", status_changed_at: new Date().toISOString(), updated_by: api.uid() }).eq("id", c.id);
      if (e3) throw e3;
    }, `Kandidat ${c.full_name} prebačen u radnike`);
    if (ok) setRows((prev) => prev.map((r) => (r.id === c.id ? { ...r, worker_id: workerId, status: "poceo" } : r)));
  };
  const copy = (text, key) => navigator.clipboard?.writeText(text).then(() => { setCopied(key); setTimeout(() => setCopied(""), 1800); }).catch(() => {});

  if (loadErr) return <Card style={{ marginTop: 8 }}><div style={{ color: S.red, fontSize: 13.5 }}>{tr("Greška: ")}{loadErr}</div></Card>;
  if (!rows) return <Card style={{ marginTop: 8 }}><div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Učitavam…")}</div></Card>;

  const inactive = new Set(["poceo", "nije_prikladan", "nedostupan"]);
  const list = rows.filter((c) => (filter === "svi" ? true : filter === "aktivni" ? !inactive.has(c.status) : c.status === filter))
    .filter((c) => !pos || (c.positions || []).includes(pos))
    // najprije oni koji najdulje čekaju na javljanje
    .sort((a, b) => (staleInfo(b)?.level === "red") - (staleInfo(a)?.level === "red") || b.created_at.localeCompare(a.created_at));
  const countBy = (st) => rows.filter((c) => c.status === st).length;
  const srcKey = (c) => c.source_tag || labelOf(HEARD, c.heard_from, LL()) || tr("bez oznake");
  const srcStats = Object.entries(rows.reduce((m, c) => {
    const k = srcKey(c); m[k] = m[k] || { n: 0, hired: 0 }; m[k].n++; if (c.status === "poceo") m[k].hired++; return m;
  }, {})).sort((a, b) => b[1].n - a[1].n);
  const adList = ads && ads.length ? ads : AD_FALLBACK;
  const ad = adList.find((a) => a.id === adId) || adList[0];
  const adBase = ad.is_cs ? `${window.location.origin}/prace` : formUrl;
  const adParams = adTag.trim() ? `z=${slugify(adTag)}` : "";
  const adLink = adParams ? `${adBase}?${adParams}` : adBase;
  const adFull = `${ad.body}\n${adLink}`;
  const saveAd = async () => {
    const e = adEdit; if (!e || !e.label.trim()) return;
    let newId = e.id;
    const ok = await api.run(async () => {
      const row = { label: e.label.trim().slice(0, 60), body: e.body.slice(0, 3000), is_cs: !!e.is_cs, updated_by: api.uid(), updated_at: new Date().toISOString() };
      if (e.id) {
        const { error } = await supabase.from("ad_texts").update(row).eq("id", e.id);
        if (error) throw error;
      } else {
        const { data: n, error } = await supabase.from("ad_texts").insert({ ...row, sort: (ads || []).length + 1 }).select().single();
        if (error) throw error;
        newId = n.id;
      }
    }, `${e.id ? "Promijenio" : "Dodao"} tekst oglasa: ${e.label.trim()}`);
    if (ok) { setAdEdit(null); await load(); setAdId(newId); }
  };
  const deleteAd = async () => {
    const e = adEdit; if (!e?.id || !window.confirm(tr("Obrisati ovaj tekst oglasa?"))) return;
    const ok = await api.run(async () => {
      const { error } = await supabase.from("ad_texts").delete().eq("id", e.id);
      if (error) throw error;
    }, `Obrisao tekst oglasa: ${e.label}`);
    if (ok) { setAdEdit(null); setAdId(""); await load(); }
  };
  const setTag = (v) => { setAdTag(v); try { localStorage.setItem("evidencija_ad_tag", v); } catch {} };
  const refLink = refWorker ? `${formUrl}?z=preporuka-${slugify(data.workers.find((w) => w.id === refWorker)?.name)}` : "";

  const chip = (active) => ({ border: `1px solid ${active ? S.green : S.line}`, background: active ? S.greenSoft : "#fff", color: active ? S.green : S.ink,
    borderRadius: 999, padding: "5px 11px", fontSize: 12.5, fontWeight: active ? 700 : 500, cursor: "pointer" });
  const tag = (bg, color) => ({ background: bg, color, borderRadius: 6, padding: "2px 7px", fontSize: 12, fontWeight: 600 });

  return (
    <div style={{ marginTop: 8 }}>
      <Card>
        <div style={{ fontWeight: 700, marginBottom: 4 }}>{tr("🔗 Link za oglase")}</div>
        <div style={{ fontSize: 13, color: S.sub, marginBottom: 8 }}>
          {tr("Stavi ga u svaku objavu u FB i Viber grupama. Kandidat ispuni upitnik na ukrajinskom (ili češkom) i prijava se pojavi ovdje.")}
        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 8 }}>
          <code style={{ flex: 1, background: S.bg, borderRadius: 8, padding: "8px 10px", fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{formUrl}</code>
          <Btn small onClick={() => copy(formUrl, "main")}>{copied === "main" ? tr("Kopirano ✓") : tr("Kopiraj")}</Btn>
        </div>
        <div style={{ fontSize: 12.5, color: S.sub }}>
          {tr("Da vidiš iz koje grupe dolaze ljudi, dodaj oznaku na kraj linka, npr.")} <code>{formUrl}?z=fb-ukrajinci-praha</code>{" "}
          {tr("ili")} <code>?z=viber-robota</code>. {tr("Češka verzija:")} <code>{window.location.origin}/prace</code>.
        </div>
      </Card>

      <Card>
        <div style={{ fontWeight: 700, marginBottom: 4 }}>{tr("📣 Tekstovi za oglase")}</div>
        <div style={{ fontSize: 13, color: S.sub, marginBottom: 10 }}>{tr("Odaberi tekst, upiši ime grupe i kopiraj. Link s oznakom grupe je već na kraju teksta.")}</div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 }}>
          {adList.map((a) => <button key={a.id} style={chip(ad.id === a.id)} onClick={() => { setAdId(a.id); setAdEdit(null); }}>{a.label}{a.is_cs ? " 🇨🇿" : ""}</button>)}
          {ads && <button style={{ ...chip(adEdit && !adEdit.id), borderStyle: "dashed" }} onClick={() => setAdEdit({ id: null, label: "", body: "", is_cs: false })}>{tr("+ Novi tekst")}</button>}
        </div>
        {adEdit ? (
          <div style={{ border: `1px solid ${S.line}`, borderRadius: 12, padding: 12, marginBottom: 10 }}>
            <Field label={tr("Naziv (vidi se samo ovdje)")}>
              <input value={adEdit.label} maxLength={60} onChange={(e) => setAdEdit({ ...adEdit, label: e.target.value })} placeholder={tr("npr. Kuhinja – vikend")} style={{ width: "100%", padding: "8px 10px", boxSizing: "border-box" }} />
            </Field>
            <Field label={tr("Tekst oglasa (link se dodaje sam na kraj)")}>
              <textarea value={adEdit.body} maxLength={3000} onChange={(e) => setAdEdit({ ...adEdit, body: e.target.value })} style={{ width: "100%", minHeight: 180, padding: "8px 10px", boxSizing: "border-box", fontSize: 14, lineHeight: 1.45 }} />
            </Field>
            <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, marginBottom: 10, cursor: "pointer" }}>
              <input type="checkbox" checked={!!adEdit.is_cs} onChange={(e) => setAdEdit({ ...adEdit, is_cs: e.target.checked })} style={{ width: 18, height: 18 }} />
              {tr("Češki tekst (link otvara češku prijavnicu)")}
            </label>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", justifyContent: "space-between" }}>
              <div style={{ display: "flex", gap: 6 }}>
                <Btn small onClick={saveAd} disabled={!adEdit.label.trim() || !adEdit.body.trim()}>{tr("Spremi")}</Btn>
                <Btn small kind="ghost" onClick={() => setAdEdit(null)}>{tr("Odustani")}</Btn>
              </div>
              {adEdit.id && <Btn small kind="danger" onClick={deleteAd}>{tr("Obriši tekst")}</Btn>}
            </div>
          </div>
        ) : (<>
        <Field label={tr("Grupa (FB ili Viber) — da znaš odakle dolaze ljudi")}>
          <input value={adTag} onChange={(e) => setTag(e.target.value)} placeholder={tr("npr. fb robota praha")} style={{ width: "100%", padding: "8px 10px", boxSizing: "border-box" }} />
        </Field>
        <div style={{ background: S.bg, borderRadius: 10, padding: "10px 12px", fontSize: 13.5, whiteSpace: "pre-wrap", lineHeight: 1.45, marginBottom: 10, wordBreak: "break-word" }}>{adFull}</div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          <Btn small onClick={() => copy(adFull, "ad")}>{copied === "ad" ? tr("Kopirano ✓") : tr("📋 Kopiraj tekst s linkom")}</Btn>
          {ads && ad.id !== "fallback" && <Btn small kind="ghost" onClick={() => setAdEdit({ id: ad.id, label: ad.label, body: ad.body, is_cs: ad.is_cs })}>{tr("✏️ Uredi")}</Btn>}
          <Btn small kind="ghost" onClick={() => copy(ad.body, "adtext")}>{copied === "adtext" ? tr("Kopirano ✓") : tr("Samo tekst")}</Btn>
          <Btn small kind="ghost" onClick={() => copy(adLink, "adlink")}>{copied === "adlink" ? tr("Kopirano ✓") : tr("Samo link")}</Btn>
        </div>
        <div style={{ fontSize: 12.5, color: S.sub, marginTop: 8 }}>{tr("Grupa ne dopušta linkove? Objavi samo tekst, a link stavi u prvi komentar.")}</div>
        </>)}
      </Card>

      <Card>
        <div style={{ fontWeight: 700, marginBottom: 4 }}>{tr("🎁 Link za preporuke")}</div>
        <div style={{ fontSize: 13, color: S.sub, marginBottom: 8 }}>{tr("Odaberi radnika i pošalji mu njegov link. Svaka prijava preko njega vidi se pod njegovim imenom, pa znaš kome ide bonus.")}</div>
        <select value={refWorker} onChange={(e) => setRefWorker(e.target.value)} style={{ width: "100%", padding: "8px 10px", marginBottom: refLink ? 8 : 0 }}>
          <option value="">{tr("— odaberi radnika —")}</option>
          {activeWorkers.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
        </select>
        {refLink && (
          <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
            <code style={{ flex: 1, background: S.bg, borderRadius: 8, padding: "8px 10px", fontSize: 12.5, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{refLink}</code>
            <Btn small onClick={() => copy(refLink, "ref")}>{copied === "ref" ? tr("Kopirano ✓") : tr("Kopiraj")}</Btn>
          </div>
        )}
      </Card>

      {rows.length > 0 && (
        <Card>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 }}>
            <button style={chip(filter === "aktivni")} onClick={() => setFilter("aktivni")}>{tr("Otvoreni")} ({rows.filter((c) => !inactive.has(c.status)).length})</button>
            {CAND_STATUS.map((s) => (
              <button key={s.id} style={chip(filter === s.id)} onClick={() => setFilter(s.id)}>{tr(s.label)} ({countBy(s.id)})</button>
            ))}
            <button style={chip(filter === "svi")} onClick={() => setFilter("svi")}>{tr("Svi")} ({rows.length})</button>
          </div>
          <select value={pos} onChange={(e) => setPos(e.target.value)} style={{ width: "100%", padding: "8px 10px" }}>
            <option value="">{tr("Sve pozicije")}</option>
            {POSITIONS.map((p) => <option key={p.id} value={p.id}>{p[LL()]} ({rows.filter((c) => (c.positions || []).includes(p.id)).length})</option>)}
          </select>
        </Card>
      )}

      {rows.length === 0 && <Card><Empty text={tr("Još nema prijava. Objavi link u grupama i prijave će se pojaviti ovdje.")} /></Card>}
      {rows.length > 0 && list.length === 0 && <Card><Empty text={tr("Nema kandidata za ovaj filter.")} /></Card>}

      {list.map((c) => {
        const st = CAND_STATUS.find((s) => s.id === c.status) || CAND_STATUS[0];
        const isOpen = open === c.id;
        const note = notes[c.id] ?? c.admin_note ?? "";
        const stale = staleInfo(c);
        const asWorker = !c.worker_id && workerByPhone[phoneKey(c.phone)];
        const repeats = (phoneCount[phoneKey(c.phone)] || 0) > 1;
        const linked = c.worker_id && data.workers.find((w) => w.id === c.worker_id);
        const msg = firstMessage(c, api.myName);
        return (
          <Card key={c.id} style={{ padding: 12, borderLeft: stale?.level === "red" ? `4px solid ${S.red}` : undefined }}>
            <div onClick={() => setOpen(isOpen ? null : c.id)} style={{ cursor: "pointer" }}>
              <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "flex-start" }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 700, fontSize: 15 }}>{c.full_name}</div>
                  <div style={{ fontSize: 12.5, color: S.sub }}>{fmtDate(c.created_at.slice(0, 10))} · {c.phone}{c.city ? ` · ${c.city}` : ""}</div>
                </div>
                <span style={{ background: st.bg, color: st.color, borderRadius: 999, padding: "3px 9px", fontSize: 12, fontWeight: 700, whiteSpace: "nowrap" }}>{tr(st.label)}</span>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginTop: 7 }}>
                {stale && <span style={tag(stale.level === "red" ? S.redSoft : S.amberSoft, stale.level === "red" ? S.red : S.amber)}>⏰ {stale.text}</span>}
                {asWorker && <span style={tag(S.blueSoft, S.blue)}>⚠ {tr("Već radnik:")} {asWorker.name}</span>}
                {repeats && <span style={tag(S.amberSoft, S.amber)}>{tr("Ponovljena prijava")}</span>}
                {(c.positions || []).map((p) => (
                  <span key={p} style={{ background: POSITIONS.find((x) => x.id === p)?.focus ? S.amberSoft : S.bg, color: S.ink, borderRadius: 6, padding: "2px 7px", fontSize: 12 }}>{labelOf(POSITIONS, p, LL())}</span>
                ))}
                {c.object_id && <span style={tag(S.greenSoft, S.green)}>🏨 {objName(c.object_id)}</span>}
                {c.source_tag && <span style={{ ...tag(S.bg, S.sub), fontWeight: 500 }}>🔗 {c.source_tag}</span>}
              </div>
            </div>

            {isOpen && (
              <div style={{ marginTop: 10, borderTop: `1px solid ${S.line}`, paddingTop: 10 }}>
                <div style={{ display: "grid", gap: 4, fontSize: 13.5, marginBottom: 10 }}>
                  {c.messenger && <div>{tr("Kontakt:")} <b>{labelOf(MESSENGERS, c.messenger, LL())}</b></div>}
                  {c.residence && <div>{tr("Boravak:")} <b>{labelOf(RESIDENCE, c.residence, LL())}</b></div>}
                  {c.start_from && <div>{tr("Može početi:")} <b>{labelOf(START, c.start_from, LL())}</b></div>}
                  {c.experience && <div>{tr("Iskustvo:")} {c.experience}</div>}
                  {(c.heard_from || c.referrer) && <div>{tr("Izvor:")} <b>{labelOf(HEARD, c.heard_from, LL()) || "—"}</b>{c.referrer ? <> · {tr("preporučio/la")} <b>{c.referrer}</b></> : null}</div>}
                  <div style={{ color: S.sub, fontSize: 12 }}>{tr("Ispunio/la na jeziku:")} {c.lang === "CS" ? "CZ" : "UA"}</div>
                </div>

                <div style={{ background: S.bg, borderRadius: 10, padding: "9px 11px", fontSize: 13, marginBottom: 8, lineHeight: 1.45 }}>
                  <div style={{ fontSize: 11.5, fontWeight: 700, color: S.sub, marginBottom: 3 }}>{tr("Prva poruka (možeš je promijeniti prije slanja)")}</div>
                  {msg}
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 12 }}>
                  {contactLinks(c, msg).map((l) => (
                    <a key={l.label} href={l.href} target={l.href.startsWith("http") ? "_blank" : undefined} rel="noreferrer"
                      onClick={() => { if (l.copy) copy(l.copy, `v-${c.id}`); if (c.status === "novi" && l.href !== `tel:${digitsPhone(c.phone)}`) setStatus(c, "kontaktiran"); }}
                      style={{ border: `1px solid ${S.line}`, borderRadius: 10, padding: "7px 11px", fontSize: 13, fontWeight: 600, color: S.ink, textDecoration: "none", background: "#fff" }}>{l.icon} {l.label}</a>
                  ))}
                  {copied === `v-${c.id}` && <span style={{ fontSize: 12.5, color: S.green, alignSelf: "center" }}>{tr("Poruka kopirana, zalijepi je u Viber.")}</span>}
                </div>

                <Field label={tr("Status")}>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                    {CAND_STATUS.map((s) => (
                      <button key={s.id} onClick={() => c.status !== s.id && setStatus(c, s.id)}
                        style={{ border: `1px solid ${c.status === s.id ? s.color : S.line}`, background: c.status === s.id ? s.bg : "#fff", color: c.status === s.id ? s.color : S.ink,
                          borderRadius: 999, padding: "6px 11px", fontSize: 12.5, fontWeight: c.status === s.id ? 700 : 500, cursor: "pointer" }}>{tr(s.label)}</button>
                    ))}
                  </div>
                </Field>
                <Field label={tr("Hotel")}>
                  <select value={c.object_id || ""} onChange={(e) => update(c, { object_id: e.target.value || null }, `Kandidat ${c.full_name}: hotel → ${objName(e.target.value) || "—"}`)} style={{ width: "100%", padding: "8px 10px" }}>
                    <option value="">{tr("— nije dodijeljen —")}</option>
                    {objects.map((o) => <option key={o.id} value={o.id}>{o.name}</option>)}
                  </select>
                </Field>
                <Field label={tr("Bilješka")}>
                  <textarea value={note} onChange={(e) => setNotes((p) => ({ ...p, [c.id]: e.target.value }))}
                    onBlur={() => note !== (c.admin_note || "") && update(c, { admin_note: note })}
                    placeholder={tr("Npr. razgovor u četvrtak 10h, govori češki…")} style={{ width: "100%", minHeight: 60, padding: "8px 10px", boxSizing: "border-box" }} />
                </Field>
                <div style={{ display: "flex", justifyContent: "space-between", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                  {linked ? <span style={{ fontSize: 13, color: S.green, fontWeight: 600 }}>✓ {tr("Radnik:")} {linked.name}</span>
                    : asWorker ? <span style={{ fontSize: 12.5, color: S.sub }}>{tr("Ovaj broj već ima radnik")} {asWorker.name}.</span>
                    : <Btn small onClick={() => toWorker(c)}>{tr("➕ Prebaci u radnike")}</Btn>}
                  {api.admin && <Btn small kind="danger" onClick={() => remove(c)}>{tr("Obriši prijavu")}</Btn>}
                </div>
              </div>
            )}
          </Card>
        );
      })}

      {srcStats.length > 0 && (
        <Card>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("📊 Odakle dolaze kandidati")}</div>
          <div style={{ display: "grid", gap: 4, fontSize: 13.5 }}>
            {srcStats.map(([k, v]) => (
              <div key={k} style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
                <span>{k}</span>
                <span style={{ color: S.sub }}>{v.n} {tr("prijava")} · <b style={{ color: v.hired ? S.green : S.sub }}>{v.hired} {tr("počelo")}</b></span>
              </div>
            ))}
          </div>
        </Card>
      )}
    </div>
  );
}
const ADMIN_MENU = [
  ["candidates", "🧑‍🍳", "Kandidati (prijave)"], ["approvals", "⏳", "Odobravanje isplata"], ["cash", "💵", "Gotovina kod koordinatora"], ["finance", "💼", "Financije (banka i porezi)"], ["unpaid", "💰", "Neisplaćeno (dugovi radnicima)"], ["vat", "🧾", "DPH"], ["users", "👥", "Korisnici i provizije"],
  ["firm", "🏢", "Podaci firme i postavke"], ["audit", "📜", "Povijest radnji"], ["logins", "🔐", "Prijave"],
  ["errors", "🐛", "Greške aplikacije"], ["trash", "🗑", "Koš (obrisano)"],
];
const adminBadge = (data, id) => id === "candidates" ? (data.newCandidates || 0) : id === "approvals" ? (data.payouts || []).filter((p) => !p.approved).length
  : id === "errors" ? (data.errorLog || []).length : id === "trash" ? data.trash.length : 0;
/* ================================================================== */
/*  NEISPLAĆENO — stare koverte i djelomične isplate                   */
/*  Tablice: unpaid_wages (koverta = radnik + mjesec), unpaid_wage_payments */
/* ================================================================== */
const monthLabel = (mo) => { if (!mo) return ""; const [y, m] = mo.split("-").map(Number); return `${String(m).padStart(2, "0")}/${y}`; };
const monthLong = (mo) => { if (!mo) return ""; const [y, m] = mo.split("-").map(Number); return `${mName(m - 1)} ${y}.`; };
// ukrajinski ima ista pravila množine (1 / 2–4 / 5+), pa se samo zamijene riječi
const PLURAL_UK = { dan: ["день", "дні", "днів"], osoba: ["особа", "особи", "осіб"], koverta: ["конверт", "конверти", "конвертів"], mjesec: ["місяць", "місяці", "місяців"], "otvoren mjesec": ["відкритий місяць", "відкриті місяці", "відкритих місяців"] };
const plural = (n, one, few, many) => { const m10 = n % 10, m100 = n % 100; const f = CUR_LANG === "UK" && PLURAL_UK[one]; if (f) [one, few, many] = f; return `${n} ${m10 === 1 && m100 !== 11 ? one : m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14) ? few : many}`; };
function UnpaidPanel({ data, api, onOpenWorker }) {
  const [rows, setRows] = useState(null);      // koverte sa stanjem (view unpaid_wages_balance)
  const [pays, setPays] = useState([]);        // pojedinačne isplate
  const [loadErr, setLoadErr] = useState("");
  const [obj, setObj] = useState(null);        // null = popis objekata, inače object_id ("none" = bez objekta)
  const [view, setView] = useState("mjeseci"); // mjeseci | osobe | ukupno
  const [selMonth, setSelMonth] = useState(null);
  const [selPerson, setSelPerson] = useState(null);
  const [showClosed, setShowClosed] = useState(false);
  const [payForm, setPayForm] = useState(null); // { key, wageId, amount, date, note }
  const [exporting, setExporting] = useState(false);
  const [newOpen, setNewOpen] = useState(false);
  const [nf, setNf] = useState({ workerId: "", objectId: "", month: prevMonthKey(curMonth()), section: "", amount: "", calc: "" });

  const load = useCallback(async () => {
    const [b, p] = await Promise.all([
      supabase.from("unpaid_wages_balance").select("*").order("month"),
      supabase.from("unpaid_wage_payments").select("*").is("deleted_at", null).order("pay_date", { ascending: true, nullsFirst: true }),
    ]);
    if (b.error || p.error) { setLoadErr((b.error || p.error).message); return; }
    setLoadErr("");
    setRows(b.data.map((r) => ({ ...r, amount: Number(r.amount), paid: Number(r.paid), remaining: Number(r.remaining) })));
    setPays(p.data.map((x) => ({ ...x, amount: Number(x.amount) })));
  }, []);
  useEffect(() => { load(); }, [load]);

  if (loadErr) return <Card style={{ marginTop: 8, color: S.red }}>{tr("Ne mogu učitati neisplaćeno:")}{" "}{loadErr}</Card>;
  if (!rows) return <Card style={{ marginTop: 8, color: S.sub }}>{tr("Učitavam…")}</Card>;

  const OPEN = 0.004;
  const isOpen = (x) => x.remaining > OPEN;
  const objKey = (r) => r.object_id || "none";
  const objName = (k) => k === "none" ? "Bez objekta" : (data.objects.find((o) => o.id === k)?.name || rows.find((r) => r.object_id === k)?.object_name || "—");
  const sum = (list, f) => list.reduce((s, x) => s + x[f], 0);
  const personKey = (r) => r.worker_id || "name:" + r.worker_name;
  const keyOf = (r) => personKey(r) + "|" + objKey(r);
  const paysOf = (wageId) => pays.filter((p) => p.unpaid_wage_id === wageId);
  const visibleRows = (list) => showClosed ? list : list.filter(isOpen);
  const totals = (list) => ({ amount: sum(list, "amount"), paid: sum(list, "paid"), remaining: sum(list, "remaining"), count: list.length });

  // grupiranje po ključu
  const groupBy = (list, keyFn, mk) => {
    const m = {};
    list.forEach((r) => { const k = keyFn(r); (m[k] ||= { key: k, ...mk(r), list: [] }).list.push(r); });
    return Object.values(m).map((g) => ({ ...g, ...totals(g.list) }));
  };

  const objList = groupBy(rows, objKey, (r) => ({ name: objName(objKey(r)) }))
    .map((o) => ({ ...o, people: new Set(o.list.filter(isOpen).map(personKey)).size }))
    .sort((a, b) => b.remaining - a.remaining || a.name.localeCompare(b.name, "hr"));
  const fr = obj ? rows.filter((r) => objKey(r) === obj) : rows;           // koverte objekta
  const months = groupBy(fr, (r) => r.month, (r) => ({ month: r.month }))
    .map((m) => ({ ...m, people: new Set(m.list.filter(isOpen).map(personKey)).size }))
    .sort((a, b) => a.month.localeCompare(b.month));
  const people = groupBy(fr, keyOf, (r) => ({ workerId: r.worker_id, name: r.worker_name, objectId: r.object_id }))
    .sort((a, b) => b.remaining - a.remaining || a.name.localeCompare(b.name, "hr"));
  const groups = Object.fromEntries(people.map((g) => [g.key, g]));
  const objT = totals(fr);

  // trenutni izbor (za naslov i Excel)
  const scopeRows = selMonth ? fr.filter((r) => r.month === selMonth) : selPerson ? fr.filter((r) => keyOf(r) === selPerson) : fr;
  const scopeLabel = !obj ? null : [objName(obj), selMonth ? monthLong(selMonth) : null, selPerson ? groups[selPerson]?.name : null].filter(Boolean).join(" · ");

  const goObj = (k) => { setObj(k); setView("mjeseci"); setSelMonth(null); setSelPerson(null); setPayForm(null); window.scrollTo({ top: 0, behavior: "smooth" }); };
  const goView = (v) => { setView(v); setSelMonth(null); setSelPerson(null); setPayForm(null); };
  const openMonth = (m) => { setView("mjeseci"); setSelPerson(null); setSelMonth(m); setPayForm(null); window.scrollTo({ top: 0, behavior: "smooth" }); };
  const openPerson = (k) => { setView("osobe"); setSelMonth(null); setSelPerson(k); setPayForm(null); window.scrollTo({ top: 0, behavior: "smooth" }); };

  // isplata: ako je odabrana koverta ide na nju, inače se raspoređuje od najstarijeg mjeseca
  const savePayment = async () => {
    const g = groups[payForm.key];
    const amount = parseNum(payForm.amount);
    if (!(amount > 0)) { alert(tr("Upiši iznos isplate.")); return; }
    let parts = [];
    if (payForm.wageId) {
      const r = g.list.find((x) => x.id === payForm.wageId);
      if (amount > r.remaining + OPEN && !confirm(tr("Iznos je veći od ostatka za {0} ({1}). Svejedno spremiti?", monthLabel(r.month), czk(r.remaining)))) return;
      parts = [{ id: r.id, amount }];
    } else {
      if (amount > g.remaining + OPEN) { alert(tr("Iznos je veći od ukupnog duga ({0}).", czk(g.remaining))); return; }
      let left = amount;
      for (const r of g.list.filter(isOpen).sort((a, b) => a.month.localeCompare(b.month))) {
        if (left <= 0) break; const a = Math.min(left, r.remaining); parts.push({ id: r.id, amount: round2(a) }); left = round2(left - a);
      }
    }
    const ok = await api.run(async () => {
      const { error } = await supabase.from("unpaid_wage_payments").insert(parts.map((x) => ({
        unpaid_wage_id: x.id, amount: x.amount, pay_date: payForm.date || null, note: payForm.note || "", created_by: api.uid(),
      })));
      if (error) throw error;
    }, `Neisplaćeno: isplatio ${czk(amount)} radniku ${g.name} (${objName(g.objectId || "none")})`);
    if (ok) { setPayForm(null); load(); }
  };
  const delPayment = async (p, name) => {
    if (!confirm(tr("Obrisati isplatu {0}?", czk(p.amount)))) return;
    const ok = await api.run(async () => {
      const { error } = await supabase.from("unpaid_wage_payments").update({ deleted_at: new Date().toISOString() }).eq("id", p.id);
      if (error) throw error;
    }, `Neisplaćeno: obrisao isplatu ${czk(p.amount)} (${name})`);
    if (ok) load();
  };
  const delWage = async (r) => {
    if (!confirm(tr("Obrisati kovertu {0} {1} ({2})?", r.worker_name, monthLabel(r.month), czk(r.amount)))) return;
    const ok = await api.run(async () => {
      const { error } = await supabase.from("unpaid_wages").update({ deleted_at: new Date().toISOString() }).eq("id", r.id);
      if (error) throw error;
    }, `Neisplaćeno: obrisao kovertu ${r.worker_name} ${r.month}`);
    if (ok) load();
  };
  const openNew = () => { setNf({ ...nf, objectId: obj && obj !== "none" ? obj : nf.objectId, month: selMonth || nf.month }); setNewOpen(!newOpen); };
  const addWage = async () => {
    const w = data.workers.find((x) => x.id === nf.workerId);
    const amount = parseNum(nf.amount);
    if (!w || !/^\d{4}-\d{2}$/.test(nf.month) || !(amount > 0)) { alert(tr("Odaberi radnika, mjesec i iznos.")); return; }
    const objectId = nf.objectId || w.objectId || null;
    const ok = await api.run(async () => {
      const { error } = await supabase.from("unpaid_wages").insert({
        worker_id: w.id, worker_name: w.name, object_id: objectId, month: nf.month, section: nf.section, amount,
        calc_note: nf.calc, created_by: api.uid(),
      });
      if (error) throw error;
    }, `Neisplaćeno: dodao kovertu ${w.name} ${nf.month} ${czk(amount)}`);
    if (ok) { setNewOpen(false); setNf({ ...nf, workerId: "", section: "", amount: "", calc: "" }); load(); }
  };

  const exportXlsx = async () => {
    if (!scopeRows.length) { alert(tr("Nema koverti za izvoz.")); return; }
    const clean = (t) => t.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^\w]+/g, "_").replace(/^_|_$/g, "");
    const suffix = scopeLabel ? "_" + clean(scopeLabel) : "";
    setExporting(true);
    try {
      await exportUnpaidXlsx({
        rows: scopeRows.map((r) => ({ ...r, objectName: objName(objKey(r)) })),
        pays, scope: scopeLabel, fileName: `Neisplaceno${suffix}_${todayISO()}.xlsx`,
      });
    } catch (e) { alert("Izvoz nije uspio: " + (e.message || e)); }
    finally { setExporting(false); }
  };

  /* ---------------- sitni dijelovi sučelja ---------------- */
  const bar = (paid, amount) => (
    <div style={{ height: 5, background: S.line, borderRadius: 99, marginTop: 9 }}>
      <div style={{ width: (amount ? Math.min(100, (paid / amount) * 100) : 0) + "%", height: "100%", background: S.green, borderRadius: 99 }} />
    </div>
  );
  const amountEl = (t, big) => (
    <div style={{ textAlign: "right", flexShrink: 0 }}>
      <div className="num" style={{ fontWeight: 800, fontSize: big ? 17 : 15.5, color: t.remaining > OPEN ? S.amber : S.green }}>{t.remaining > OPEN ? czk(t.remaining) : tr("✓ isplaćeno")}</div>
      <div className="num" style={{ fontSize: 11.5, color: S.sub }}>{tr("od")}{" "}{czk(t.amount)}</div>
    </div>
  );
  const rowCard = ({ key, title, sub, t, onClick }) => (
    <Card key={key} onClick={onClick} style={{ cursor: onClick ? "pointer" : "default", padding: "12px 14px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
        <div style={{ minWidth: 0 }}>
          <div style={{ fontWeight: 800, fontSize: 15.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{title}</div>
          {sub && <div style={{ fontSize: 12.5, color: S.sub, marginTop: 1 }}>{sub}</div>}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>{amountEl(t)}{onClick && <span style={{ color: S.sub, fontSize: 18 }}>›</span>}</div>
      </div>
      {bar(t.paid, t.amount)}
    </Card>
  );
  const summary = ({ label, t, sub }) => (
    <Card style={{ background: S.amberSoft, borderColor: "#EBD9AE" }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: S.amber, textTransform: "uppercase", letterSpacing: 0.6 }}>{label}</div>
      <div className="num" style={{ fontSize: 30, fontWeight: 800, color: S.ink, margin: "2px 0 6px" }}>{czk(t.remaining)}</div>
      <div style={{ display: "flex", gap: 14, flexWrap: "wrap", fontSize: 12.5, color: S.sub }} className="num">
        <span>{tr("Na kovertama")}{" "}<b style={{ color: S.ink }}>{czk(t.amount)}</b></span>
        <span>{tr("Isplaćeno")}{" "}<b style={{ color: S.green }}>{czk(t.paid)}</b></span>
        {sub && <span>{sub}</span>}
      </div>
      {bar(t.paid, t.amount)}
      <div style={{ display: "flex", gap: 8, marginTop: 12, flexWrap: "wrap" }}>
        <Btn small kind="excel" onClick={exportXlsx} disabled={exporting}>{exporting ? tr("Pripremam…") : tr("⬇ Excel")}</Btn>
        <Btn small kind="ghost" onClick={openNew}>{tr("＋ Nova koverta")}</Btn>
        <Btn small kind="ghost" onClick={() => setShowClosed(!showClosed)}>{showClosed ? tr("Sakrij isplaćene") : tr("Prikaži isplaćene")}</Btn>
      </div>
    </Card>
  );
  const crumbs = (items) => (
    <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap", marginBottom: 10, fontSize: 13.5 }}>
      {items.map(([label, fn], i) => (
        <span key={i} style={{ display: "flex", alignItems: "center", gap: 6 }}>
          {i > 0 && <span style={{ color: S.sub }}>›</span>}
          {fn ? <button onClick={fn} style={{ background: "none", border: "none", padding: 0, color: S.blue, fontWeight: 600, cursor: "pointer", fontSize: 13.5 }}>{label}</button>
            : <b>{label}</b>}
        </span>
      ))}
    </div>
  );
  const tabs = () => (
    <div style={{ display: "flex", background: "#E9ECE6", borderRadius: 12, padding: 3, marginBottom: 12 }}>
      {[["mjeseci", "Po mjesecima"], ["osobe", "Po osobama"], ["ukupno", "Ukupno"]].map(([id, label]) => {
        const on = view === id && !selMonth && !selPerson;
        const active = view === id;
        return (
          <button key={id} onClick={() => goView(id)} style={{ flex: 1, border: "none", borderRadius: 10, padding: "9px 6px", cursor: "pointer", fontSize: 13.5,
            fontWeight: active ? 700 : 500, background: active ? "#fff" : "transparent", color: active ? S.green : S.sub,
            boxShadow: active ? "0 1px 3px rgba(0,0,0,.08)" : "none", opacity: on || active ? 1 : 0.9 }}>{label}</button>
        );
      })}
    </div>
  );
  const payFormEl = (g) => (
    <div style={{ marginTop: 10, background: "#fff", border: `1px solid ${S.line}`, borderRadius: 12, padding: 12 }}>
      <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("💶 Isplata —")}{" "}{g.name}</div>
      <Field label={tr("Za koji mjesec")}>
        <select value={payForm.wageId} onChange={(e) => setPayForm({ ...payForm, wageId: e.target.value })}>
          <option value="">{tr("Automatski — od najstarijeg mjeseca")}</option>
          {g.list.filter(isOpen).sort((a, b) => a.month.localeCompare(b.month)).map((r) => (
            <option key={r.id} value={r.id}>{monthLabel(r.month)} {tr("— ostaje")}{" "}{czk(r.remaining)}</option>
          ))}
        </select>
      </Field>
      <div style={{ display: "flex", gap: 10 }}>
        <div style={{ flex: 1 }}><Field label={tr("Iznos (Kč)")}><input inputMode="decimal" autoFocus value={payForm.amount} onChange={(e) => setPayForm({ ...payForm, amount: e.target.value })} /></Field></div>
        <div style={{ flex: 1 }}><Field label={tr("Datum")}><input type="date" value={payForm.date} onChange={(e) => setPayForm({ ...payForm, date: e.target.value })} /></Field></div>
      </div>
      <Field label={tr("Napomena (nije obavezno)")}><input value={payForm.note} onChange={(e) => setPayForm({ ...payForm, note: e.target.value })} placeholder={tr("npr. gotovina, banka")} /></Field>
      <div style={{ display: "flex", gap: 8 }}><Btn small onClick={savePayment}>{tr("Spremi isplatu")}</Btn><Btn small kind="ghost" onClick={() => setPayForm(null)}>{tr("Odustani")}</Btn></div>
    </div>
  );
  // jedna koverta s isplatama i gumbima; label = mjesec ili ime osobe
  const envelope = (r, label) => {
    const g = groups[keyOf(r)];
    return (
      <div key={r.id} style={{ padding: "11px 0", borderBottom: `1px solid ${S.line}` }}>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "baseline" }}>
          <div style={{ fontWeight: 700, fontSize: 14.5, minWidth: 0 }}>{label} {r.section && <span style={{ fontWeight: 500, color: S.sub, fontSize: 12.5 }}>· {r.section}</span>}</div>
          <div className="num" style={{ fontSize: 13.5, whiteSpace: "nowrap" }}>
            <span style={{ color: S.sub }}>{czk(r.amount)}</span>{" → "}
            <b style={{ color: isOpen(r) ? S.amber : S.green }}>{isOpen(r) ? czk(r.remaining) : "✓"}</b>
          </div>
        </div>
        {r.calc_note && <div style={{ fontSize: 11.5, color: S.sub, marginTop: 2 }}>{r.calc_note}</div>}
        {r.note && <div style={{ fontSize: 11.5, color: S.amber, marginTop: 2 }}>ⓘ {r.note}</div>}
        {paysOf(r.id).map((p) => (
          <div key={p.id} className="num" style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, marginTop: 4, color: S.green }}>
            <span style={{ flex: 1 }}>− {czk(p.amount)} · {p.pay_date ? fmtDate(p.pay_date) : tr("bez datuma")}{p.note ? " · " + p.note : ""}</span>
            <button onClick={() => delPayment(p, r.worker_name)} title={tr("Obriši isplatu")} style={{ background: "none", border: "none", color: S.red, cursor: "pointer", fontSize: 14 }}>✕</button>
          </div>
        ))}
        <div style={{ display: "flex", gap: 6, marginTop: 7, flexWrap: "wrap" }}>
          {isOpen(r) && <Btn small kind="ghost" onClick={() => setPayForm({ key: keyOf(r), wageId: r.id, amount: String(r.remaining), date: todayISO(), note: "", at: r.id })}>{tr("💶 Isplati")}</Btn>}
          <Btn small kind="danger" onClick={() => delWage(r)}>{tr("Obriši")}</Btn>
        </div>
        {payForm && payForm.at === r.id && g && payFormEl(g)}
      </div>
    );
  };

  const workerOptions = sortedWorkers(data.workers);
  const objectOptions = sortedObjects(data.objects.filter((o) => !o.archived || o.id === nf.objectId));
  const newForm = newOpen && (
    <Card>
      <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("＋ Nova neisplaćena koverta")}</div>
      <Field label={tr("Objekt")}>
        <select value={nf.objectId} onChange={(e) => setNf({ ...nf, objectId: e.target.value })}>
          <option value="">{tr("— objekt radnika —")}</option>
          {objectOptions.map((o) => <option key={o.id} value={o.id}>{o.name}</option>)}
        </select>
      </Field>
      <Field label={tr("Radnik")}>
        <select value={nf.workerId} onChange={(e) => setNf({ ...nf, workerId: e.target.value })}>
          <option value="">{tr("— odaberi —")}</option>
          {workerOptions.filter((w) => !nf.objectId || w.objectId === nf.objectId || w.archived).map((w) => <option key={w.id} value={w.id}>{w.name}{w.archived ? tr(" (arhiviran)") : ""}</option>)}
        </select>
      </Field>
      <div style={{ display: "flex", gap: 10 }}>
        <div style={{ flex: 1 }}><Field label={tr("Mjesec")}><input type="month" value={nf.month} onChange={(e) => setNf({ ...nf, month: e.target.value })} /></Field></div>
        <div style={{ flex: 1 }}><Field label={tr("Iznos (Kč)")}><input inputMode="decimal" value={nf.amount} onChange={(e) => setNf({ ...nf, amount: e.target.value })} /></Field></div>
      </div>
      <Field label={tr("Odjel (npr. bílá, studená)")}><input value={nf.section} onChange={(e) => setNf({ ...nf, section: e.target.value })} /></Field>
      <Field label={tr("Izračun s koverte (nije obavezno)")}><input value={nf.calc} onChange={(e) => setNf({ ...nf, calc: e.target.value })} placeholder={tr("npr. 250 h × 150 − soc.-zdr. 5 500 − banka 19 011")} /></Field>
      <div style={{ display: "flex", gap: 8 }}><Btn small onClick={addWage}>{tr("Spremi kovertu")}</Btn><Btn small kind="ghost" onClick={() => setNewOpen(false)}>{tr("Odustani")}</Btn></div>
    </Card>
  );

  /* ---------------- 1) popis objekata ---------------- */
  if (!obj) {
    const allT = totals(rows);
    const shown = showClosed ? objList : objList.filter(isOpen);
    return (
      <div style={{ marginTop: 8 }}>
        {summary({ label: "Ukupno dugujem — svi objekti", t: allT, sub: `${objList.filter(isOpen).length} objekata` })}
        {newForm}
        <div style={{ fontSize: 12.5, fontWeight: 700, color: S.sub, textTransform: "uppercase", letterSpacing: 0.6, margin: "4px 2px 8px" }}>{tr("Odaberi objekt")}</div>
        {shown.length === 0 && <Empty text={tr("Nema neisplaćenih koverti. 🎉")} />}
        {shown.map((o) => rowCard({ key: o.key, title: `🏨 ${o.name}`, sub: `${plural(o.people, "osoba", "osobe", "osoba")} s dugom · ${plural(o.list.filter(isOpen).length, "koverta", "koverte", "koverti")}`, t: o, onClick: () => goObj(o.key) }))}
      </div>
    );
  }

  const crumbsBase = [["Svi objekti", () => goObj(null)]];

  /* ---------------- 2a) jedan mjesec ---------------- */
  if (selMonth) {
    const m = months.find((x) => x.month === selMonth);
    const list = visibleRows(m ? m.list : []).sort((a, b) => b.remaining - a.remaining || a.worker_name.localeCompare(b.worker_name, "hr"));
    const t = totals(m ? m.list : []);
    return (
      <div style={{ marginTop: 8 }}>
        {crumbs([...crumbsBase, [objName(obj), () => goView("mjeseci")], [monthLong(selMonth)]])}
        {summary({ label: `${objName(obj)} — ${monthLong(selMonth)}`, t, sub: plural(new Set((m?.list || []).filter(isOpen).map(personKey)).size, "osoba", "osobe", "osoba") })}
        {newForm}
        <Card style={{ padding: "2px 14px 6px" }}>
          {list.length === 0 && <Empty text={tr("U ovom mjesecu sve je isplaćeno. 🎉")} />}
          {list.map((r) => envelope(r,
            <button onClick={() => openPerson(keyOf(r))} style={{ background: "none", border: "none", padding: 0, fontWeight: 700, fontSize: 14.5, color: S.blue, cursor: "pointer", textAlign: "left" }}>{r.worker_name}</button>))}
        </Card>
      </div>
    );
  }

  /* ---------------- 2b) jedna osoba ---------------- */
  const pg = selPerson ? groups[selPerson] : null;
  if (pg) {
    const g = pg;
    const list = visibleRows(g.list).sort((a, b) => a.month.localeCompare(b.month));
    return (
      <div style={{ marginTop: 8 }}>
        {crumbs([...crumbsBase, [objName(obj), () => goView("osobe")], [g.name]])}
        {summary({ label: g.name, t: g, sub: plural(g.list.filter(isOpen).length, "otvoren mjesec", "otvorena mjeseca", "otvorenih mjeseci") })}
        {newForm}
        <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
          {isOpen(g) && <Btn small onClick={() => setPayForm({ key: g.key, wageId: "", amount: "", date: todayISO(), note: "", at: "person" })}>{tr("💶 Dodaj isplatu")}</Btn>}
          {g.workerId && onOpenWorker && <Btn small kind="ghost" onClick={() => onOpenWorker(g.workerId)}>{tr("Otvori radnika")}</Btn>}
        </div>
        {payForm && payForm.at === "person" && payFormEl(g)}
        <Card style={{ padding: "2px 14px 6px", marginTop: 10 }}>
          {list.length === 0 && <Empty text={tr("Sve je isplaćeno. 🎉")} />}
          {list.map((r) => envelope(r, monthLong(r.month)))}
        </Card>
      </div>
    );
  }

  /* ---------------- 2c) objekt: kartice ---------------- */
  const shownMonths = showClosed ? months : months.filter(isOpen);
  const shownPeople = showClosed ? people : people.filter(isOpen);
  return (
    <div style={{ marginTop: 8 }}>
      {crumbs([...crumbsBase, [objName(obj)]])}
      {tabs()}
      {summary({ label: `Dugujem — ${objName(obj)}`, t: objT, sub: `${plural(people.filter(isOpen).length, "osoba", "osobe", "osoba")} · ${plural(months.filter(isOpen).length, "mjesec", "mjeseca", "mjeseci")}` })}
      {newForm}

      {view === "mjeseci" && (
        <>
          {shownMonths.length === 0 && <Empty text={tr("Nema neisplaćenih mjeseci. 🎉")} />}
          {shownMonths.map((m) => rowCard({ key: m.month, title: monthLong(m.month), sub: `${plural(m.people, "osoba", "osobe", "osoba")} · ${plural(m.list.filter(isOpen).length, "koverta", "koverte", "koverti")}`, t: m, onClick: () => openMonth(m.month) }))}
        </>
      )}

      {view === "osobe" && (
        <>
          {shownPeople.length === 0 && <Empty text={tr("Nema neisplaćenih osoba. 🎉")} />}
          {shownPeople.map((g) => rowCard({ key: g.key, title: g.name, sub: g.list.filter(isOpen).sort((a, b) => a.month.localeCompare(b.month)).map((r) => monthLabel(r.month)).join(", ") || "sve isplaćeno", t: g, onClick: () => openPerson(g.key) }))}
        </>
      )}

      {view === "ukupno" && (
        <Card style={{ padding: 0, overflow: "hidden" }}>
          {[["Po mjesecima", shownMonths.map((m) => ({ k: m.month, label: monthLong(m.month), t: m, go: () => openMonth(m.month) }))],
            ["Po osobama", shownPeople.map((g) => ({ k: g.key, label: g.name, t: g, go: () => openPerson(g.key) }))]].map(([title, items]) => (
            <div key={title}>
              <div style={{ display: "flex", justifyContent: "space-between", padding: "10px 14px", background: S.greenSoft, fontSize: 12, fontWeight: 700, color: S.green, textTransform: "uppercase", letterSpacing: 0.5 }}>
                <span>{title}</span><span>{tr("Ostaje")}</span>
              </div>
              {items.map((x) => (
                <div key={x.k} onClick={x.go} style={{ display: "flex", justifyContent: "space-between", gap: 10, padding: "9px 14px", borderBottom: `1px solid ${S.line}`, cursor: "pointer", fontSize: 14 }}>
                  <span style={{ minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{x.label}</span>
                  <span className="num" style={{ fontWeight: 700, color: isOpen(x.t) ? S.amber : S.green, whiteSpace: "nowrap" }}>{czk(x.t.remaining)}</span>
                </div>
              ))}
              <div style={{ display: "flex", justifyContent: "space-between", padding: "10px 14px", fontWeight: 800, fontSize: 14, borderBottom: `1px solid ${S.line}` }} className="num">
                <span>{tr("Ukupno")}</span><span>{czk(objT.remaining)}</span>
              </div>
            </div>
          ))}
        </Card>
      )}
    </div>
  );
}

/* ================================================================== */
/*  FINANCIJE — bankovni izvodi, mjesečni pregled, procjena poreza     */
/*  Tablice: bank_imports, bank_transactions, bank_rules (samo admin)  */
/* ================================================================== */
function FinancePanel({ data, api }) {
  const rate = Number(data.settings?.czk_rate) || 25;
  const thisYear = new Date().getFullYear();
  const [year, setYear] = useState(thisYear);
  const [txs, setTxs] = useState(null);
  const [imports, setImports] = useState([]);
  const [rules, setRules] = useState([]);
  const [loadErr, setLoadErr] = useState("");
  const [view, setView] = useState("pregled"); // pregled | obveze | transakcije | uvoz | pravila
  const [fMonth, setFMonth] = useState("");
  const [fCat, setFCat] = useState("all");
  const [q, setQ] = useState("");
  const [limit, setLimit] = useState(150);
  const [remember, setRemember] = useState(true);
  const [revBasis, setRevBasis] = useState("banka"); // banka | fakture
  const [imp, setImp] = useState(null); // { fileName, rows, header, map }
  const [manual, setManual] = useState(null);
  const [ruleForm, setRuleForm] = useState({ pattern: "", category: "" });
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");

  const load = useCallback(async () => {
    try {
      const all = [];
      for (let from = 0; ; from += 1000) {
        const { data: d, error } = await supabase.from("bank_transactions").select("*")
          .gte("tx_date", `${year}-01-01`).lte("tx_date", `${year}-12-31`)
          .order("tx_date", { ascending: false }).order("id").range(from, from + 999);
        if (error) throw error;
        all.push(...d);
        if (d.length < 1000) break;
      }
      const [im, ru] = await Promise.all([
        supabase.from("bank_imports").select("*").order("created_at", { ascending: false }).limit(100),
        supabase.from("bank_rules").select("*").order("created_at"),
      ]);
      if (im.error || ru.error) throw im.error || ru.error;
      setTxs(all.map((t) => ({ ...t, amount: Number(t.amount), vat_rate: t.vat_rate == null ? null : Number(t.vat_rate) })));
      setImports(im.data); setRules(ru.data); setLoadErr("");
    } catch (e) { setLoadErr(e.message || String(e)); }
  }, [year]);
  useEffect(() => { load(); }, [load]);

  const audit = (action) => supabase.from("audit_log").insert({ user_id: api.uid(), user_name: data.profiles.find((p) => p.id === api.uid())?.name || "", action }).then(() => {});
  const run = async (fn) => {
    setBusy(true); setMsg("");
    try { await fn(); await load(); }
    catch (e) { setMsg("Greška: " + (e.message || e)); }
    finally { setBusy(false); }
  };

  if (loadErr) return (
    <Card style={{ marginTop: 8, color: S.red, fontSize: 13.5 }}>
      Ne mogu učitati financije: {loadErr}
      <div style={{ color: S.sub, marginTop: 6 }}>Ako tablice još ne postoje, u Supabase → SQL Editor pokreni zadnji dio datoteke <b>supabase/schema.sql</b> (odjeljak FINANCIJE).</div>
    </Card>
  );
  if (!txs) return <Card style={{ marginTop: 8, color: S.sub }}>Učitavam…</Card>;

  /* ---------- izračuni ---------- */
  const { months, byCat } = finYear(txs, year, rate);
  const invNet = Array(12).fill(0);
  (data.objectInvoices || []).forEach((i) => {
    const mo = invoiceTaxMonth(i);
    if (mo?.startsWith(String(year))) invNet[Number(mo.slice(5, 7)) - 1] += toCzk(i.amount - (i.vatAmount || 0), i.currency, rate);
  });
  // broj mjeseci s podacima (ne zadnji mjesec) — da projekcija radi i kad izvode počneš dodavati usred godine
  const covered = months.filter((m, i) => m.count || (revBasis === "fakture" && invNet[i])).length;
  const incomeOf = (i) => (revBasis === "fakture" ? invNet[i] : months[i].income);
  const sumM = (f) => months.reduce((s, m) => s + m[f], 0);
  const income = months.reduce((s, _m, i) => s + incomeOf(i), 0);
  const cost = sumM("cost");
  const est = taxEstimate({ income, cost, monthsCovered: covered });
  const citPaid = sumM("taxPaid");
  const uncat = txs.filter((t) => !t.category).length;
  const autoCount = txs.filter((t) => t.category && t.category_auto).length;
  const nowMo = curMonth();
  const vatMonths = months.map((m) => m.month).filter((mo) => mo <= nowMo);
  const vatRows = vatMonths.map((mo) => ({ mo, v: vatSummary(data, mo) }));
  const vatOwed = vatRows.reduce((s, r) => s + r.v.payable, 0);
  const vatPaid = sumM("vatPaid");
  const nextVat = vatRows.find((r) => r.v.payable > 0 && daysUntil(r.v.due) >= 0);
  const lev = payrollLevies({ txs, reminders: data.reminders || [], year, rate });
  const envelopeSocZdr = (data.payrollNotes || []).filter((n) => n.month?.startsWith(`${year}-`))
    .reduce((s, n) => s + (n.currency === "CZK" ? n.amount : n.sourceCzk || n.amount * rate), 0);
  const adv = citAdvances(est.yCit);
  const Kc = (n) => czk(Math.round(n));

  const filtered = txs.filter((t) => (!fMonth || t.tx_date.startsWith(fMonth))
    && (fCat === "all" || (fCat === "none" ? !t.category : fCat === "auto" ? t.category && t.category_auto : t.category === fCat))
    && (!q || fold([t.counterparty, t.counter_account, t.vs, t.message, t.note, String(t.amount)].join(" ")).includes(fold(q))));

  /* ---------- akcije ---------- */
  const patchTx = async (ids, patch) => {
    setTxs((list) => list.map((t) => (ids.includes(t.id) ? { ...t, ...patch } : t)));
    for (let i = 0; i < ids.length; i += 200) {
      const { error } = await supabase.from("bank_transactions").update(patch).in("id", ids.slice(i, i + 200));
      if (error) { setMsg("Greška: " + error.message); load(); return; }
    }
  };
  const setCat = async (tx, category) => {
    await patchTx([tx.id], { category, category_auto: false });
    const pattern = (tx.counterparty || tx.counter_account || "").trim();
    if (!remember || !category || !pattern) return;
    const existing = rules.find((r) => fold(r.pattern) === fold(pattern));
    if (existing?.category === category) return;
    const { error } = existing
      ? await supabase.from("bank_rules").update({ category }).eq("id", existing.id)
      : await supabase.from("bank_rules").insert({ pattern, category, created_by: api.uid() });
    if (error) { setMsg("Greška pri spremanju pravila: " + error.message); return; }
    const others = txs.filter((t) => t.id !== tx.id && (!t.category || t.category_auto) && t.category !== category && ruleMatches(txText(t), pattern)).map((t) => t.id);
    if (others.length) await patchTx(others, { category, category_auto: true });
    setMsg(`Zapamćeno: "${pattern}" → ${catLabel(category)}${others.length ? ` (razvrstano još ${others.length})` : ""}`);
    const { data: ru } = await supabase.from("bank_rules").select("*").order("created_at");
    if (ru) setRules(ru);
  };
  const applyRules = () => run(async () => {
    const groups = {};
    txs.filter((t) => !t.category || t.category_auto).forEach((t) => {
      const c = suggestCategory(t, rules);
      if (c && c !== t.category) (groups[c] ||= []).push(t.id);
    });
    let n = 0;
    for (const [c, ids] of Object.entries(groups)) { await patchTx(ids, { category: c, category_auto: true }); n += ids.length; }
    setMsg(n ? `Razvrstano po pravilima: ${n}` : "Nema novih prijedloga.");
  });
  const confirmAuto = () => run(async () => {
    const ids = filtered.filter((t) => t.category && t.category_auto).map((t) => t.id);
    await patchTx(ids, { category_auto: false });
    setMsg(`Potvrđeno: ${ids.length}`);
  });

  const onFile = async (file) => {
    if (!file) return;
    setMsg("");
    try {
      const buf = await file.arrayBuffer();
      let rows;
      if (/\.xlsx?$/i.test(file.name)) {
        const wb = XLSX.read(buf, { type: "array", cellDates: true });
        rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: "" });
      } else rows = parseCsv(decodeText(buf));
      const h = findHeader(rows);
      setImp({ fileName: file.name, rows, header: Math.max(h, 0), map: guessMapping(rows[Math.max(h, 0)] || []), notFound: h < 0 });
    } catch (e) { setMsg("Ne mogu pročitati datoteku: " + (e.message || e)); }
  };
  const parsed = imp ? rowsToTxs(imp.rows, imp.header, imp.map) : null;
  const doImport = () => run(async () => {
    const list = parsed.txs;
    if (!list.length) throw new Error("Nema transakcija za uvoz — provjeri stupce.");
    const dates = list.map((t) => t.tx_date).sort();
    const { data: im, error } = await supabase.from("bank_imports")
      .insert({ file_name: imp.fileName, date_from: dates[0], date_to: dates[dates.length - 1], created_by: api.uid() }).select().single();
    if (error) throw error;
    let inserted = 0;
    for (let i = 0; i < list.length; i += 500) {
      const chunk = list.slice(i, i + 500).map((t) => {
        const c = suggestCategory(t, rules);
        return { ...t, import_id: im.id, category: c, category_auto: !!c, created_by: api.uid() };
      });
      const { data: ins, error: e2 } = await supabase.from("bank_transactions").upsert(chunk, { onConflict: "hash", ignoreDuplicates: true }).select("id");
      if (e2) throw e2;
      inserted += ins.length;
    }
    if (inserted) await supabase.from("bank_imports").update({ row_count: inserted }).eq("id", im.id);
    else await supabase.from("bank_imports").delete().eq("id", im.id);
    await audit(`Uvezao bankovni izvod ${imp.fileName}: ${inserted} novih transakcija (${fmtDate(dates[0])} – ${fmtDate(dates[dates.length - 1])})`);
    setMsg(`Uvezeno ${inserted} novih transakcija${list.length - inserted ? `, ${list.length - inserted} već postoji pa je preskočeno` : ""}. Provjeri prijedloge kategorija.`);
    setImp(null);
    setYear(Number(dates[dates.length - 1].slice(0, 4)));
    setFMonth(""); setFCat(inserted ? "auto" : "all"); setView("transakcije");
  });
  const delImport = (im) => window.confirm(`Obrisati uvoz "${im.file_name}" i svih ${im.row_count} transakcija iz njega?`) && run(async () => {
    const { error } = await supabase.from("bank_imports").delete().eq("id", im.id);
    if (error) throw error;
    await audit(`Obrisao bankovni izvod ${im.file_name} (${im.row_count} transakcija)`);
  });
  const delTx = (t) => window.confirm(`Obrisati transakciju ${fmtDate(t.tx_date)} ${money(t.amount, t.currency)}?`) && run(async () => {
    const { error } = await supabase.from("bank_transactions").delete().eq("id", t.id);
    if (error) throw error;
  });
  const saveManual = () => run(async () => {
    const amount = round2(parseNum(manual.amount) || 0) * (manual.sign === "-" ? -1 : 1);
    if (!manual.date || !amount) throw new Error("Upiši datum i iznos.");
    const { error } = await supabase.from("bank_transactions").insert({
      tx_date: manual.date, amount, currency: manual.currency, counterparty: manual.counterparty.trim(), message: "Ručni upis",
      category: manual.category, category_auto: false, hash: `manual|${Date.now()}|${Math.random().toString(36).slice(2)}`, created_by: api.uid(),
    });
    if (error) throw error;
    await audit(`Ručno upisao financijsku stavku: ${manual.counterparty} ${money(amount, manual.currency)} (${fmtDate(manual.date)})`);
    setManual(null);
  });
  const addRule = () => run(async () => {
    if (!ruleForm.pattern.trim() || !ruleForm.category) return;
    const { error } = await supabase.from("bank_rules").insert({ pattern: ruleForm.pattern.trim(), category: ruleForm.category, created_by: api.uid() });
    if (error) throw error;
    setRuleForm({ pattern: "", category: "" });
  });
  const delRule = (r) => run(async () => {
    const { error } = await supabase.from("bank_rules").delete().eq("id", r.id);
    if (error) throw error;
  });

  /* ---------- prikaz ---------- */
  const chip = (on) => ({ padding: "6px 11px", borderRadius: 999, fontWeight: 700, fontSize: 12.5, cursor: "pointer",
    background: on ? S.green : "#fff", color: on ? "#fff" : S.sub, border: `1px solid ${on ? S.green : S.line}` });
  const row = { display: "flex", justifyContent: "space-between", gap: 8, padding: "5px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13 };
  const sel = { width: "auto", padding: "6px 8px", fontSize: 13 };
  const catSelect = (value, onChange, style) => (
    <select value={value} onChange={(e) => onChange(e.target.value)} style={{ ...sel, ...style }}>
      <option value="">❔ Nerazvrstano</option>
      {FIN_CATS.map(([k, label]) => <option key={k} value={k}>{label}</option>)}
    </select>
  );
  const monthName = (mo) => `${mName(Number(mo.slice(5, 7)) - 1)} ${mo.slice(0, 4)}.`;

  return (
    <div style={{ marginTop: 8 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {[["pregled", "📊 Pregled"], ["obveze", "🧮 Kraj godine"], ["transakcije", "📋 Transakcije"], ["uvoz", "📥 Uvoz izvoda"], ["pravila", "⚙️ Pravila"]].map(([id, label]) => (
            <button key={id} onClick={() => setView(id)} style={chip(view === id)}>{label}</button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
          <button onClick={() => setYear(year - 1)} style={{ background: "#fff", border: `1px solid ${S.line}`, borderRadius: 8, width: 26, height: 26, cursor: "pointer", fontWeight: 700 }}>‹</button>
          <b style={{ padding: "0 4px", fontSize: 14 }}>{year}.</b>
          <button onClick={() => setYear(year + 1)} disabled={year >= thisYear} style={{ background: "#fff", border: `1px solid ${S.line}`, borderRadius: 8, width: 26, height: 26, cursor: "pointer", fontWeight: 700, opacity: year >= thisYear ? 0.4 : 1 }}>›</button>
        </div>
      </div>
      {busy && <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 8 }}>Spremam…</div>}
      {msg && <div style={{ background: msg.startsWith("Greška") || msg.startsWith("Ne mogu") ? S.redSoft : S.greenSoft, color: msg.startsWith("Greška") || msg.startsWith("Ne mogu") ? S.red : S.green, borderRadius: 10, padding: "8px 12px", marginBottom: 10, fontSize: 13, fontWeight: 600 }}>
        {msg} <button onClick={() => setMsg("")} style={{ background: "none", border: "none", color: "inherit", fontWeight: 800, cursor: "pointer", float: "right" }}>✕</button>
      </div>}

      {view === "pregled" && (
        <>
          {txs.length === 0 && (
            <Card style={{ background: S.blueSoft, borderColor: S.blueSoft }}>
              <div style={{ fontWeight: 700, marginBottom: 4 }}>Još nema transakcija za {year}.</div>
              <div style={{ fontSize: 13.5, marginBottom: 8 }}>Izvezi izvod iz internet bankarstva kao CSV (ili Excel) i učitaj ga ovdje. Svaki mjesec samo dodaš novi — duplikati se preskaču sami.</div>
              <Btn small onClick={() => setView("uvoz")}>📥 Uvezi izvod</Btn>
            </Card>
          )}
          {(uncat > 0 || autoCount > 0) && (
            <Card style={{ background: S.amberSoft, borderColor: S.amberSoft, fontSize: 13.5 }}>
              {uncat > 0 && <div>❔ <b>{uncat}</b> transakcija nije razvrstano — isplate bez kategorije se ne odbijaju kao trošak.</div>}
              {autoCount > 0 && <div>🤖 <b>{autoCount}</b> kategorija je predloženo automatski — provjeri ih.</div>}
              <Btn small kind="ghost" style={{ marginTop: 6 }} onClick={() => { setFMonth(""); setFCat(uncat ? "none" : "auto"); setView("transakcije"); }}>Pregledaj</Btn>
            </Card>
          )}

          <Card>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
              <div style={{ fontWeight: 700 }}>💼 {year}. do sada (bez DPH)</div>
              <div style={{ display: "flex", gap: 4 }}>
                <button onClick={() => setRevBasis("banka")} style={chip(revBasis === "banka")}>prihod po banci</button>
                <button onClick={() => setRevBasis("fakture")} style={chip(revBasis === "fakture")}>po fakturama</button>
              </div>
            </div>
            <div style={row}><span>Prihodi</span><b className="num" style={{ color: S.green }}>{Kc(income)}</b></div>
            <div style={row}><span>Porezno priznati troškovi</span><b className="num" style={{ color: S.red }}>− {Kc(cost)}</b></div>
            <div style={{ ...row, fontSize: 15, borderBottom: "none" }}><b>Dobit</b><b className="num" style={{ color: est.profit >= 0 ? S.green : S.red }}>{Kc(est.profit)}</b></div>
            {sumM("owner") > 0 && <div style={{ fontSize: 12.5, color: S.sub }}>Isplaćeno vlasniku ove godine: {Kc(sumM("owner"))} (nije trošak)</div>}
            <div style={{ fontSize: 11.5, color: S.sub, marginTop: 4 }}>
              {revBasis === "banka" ? "Prihod = uplate na račun umanjene za DPH (prema stopi kategorije)." : "Prihod = izdane fakture (bez DPH) iz kartice Objekti, po DUZP-u."} Troškovi = isplate s izvoda bez DPH-a. Plaće isplaćene u gotovini nisu na izvodu — dodaj ih ručno u Transakcijama ako idu u knjige.
            </div>
          </Card>

          <Card>
            <div style={{ fontWeight: 700, marginBottom: 8 }}>🏛 Procjena poreza na dobit ({Math.round(CZ_CIT * 100)} %)</div>
            <div style={row}><span>Do sada: osnovica {Kc(est.base)}</span><b className="num">{Kc(est.cit)}</b></div>
            {est.factor > 1 && <div style={row}><span>Ako nastavi ovako cijelu godinu ({covered} mj. s podacima → 12): dobit {Kc(est.yProfit)}</span><b className="num">{Kc(est.yCit)}</b></div>}
            {citPaid !== 0 && <div style={row}><span>Već plaćene zálohe ove godine</span><b className="num" style={{ color: S.green }}>− {Kc(citPaid)}</b></div>}
            <div style={{ ...row, fontSize: 15, borderBottom: "none", background: S.amberSoft, borderRadius: 10, padding: "8px 10px", marginTop: 6 }}>
              <b>Za platiti za {year}. (procjena)</b><b className="num" style={{ color: S.amber }}>{Kc(Math.max(0, est.yCit - citPaid))}</b>
            </div>
            <div style={{ fontSize: 11.5, color: S.sub, marginTop: 6 }}>
              Prijava (přiznání k DPPO) za {year}. predaje se do 1. 4. {year + 1}. na papiru, elektronički mjesec dana kasnije, a preko poreznog savjetnika do 1. 7. {year + 1}. Osnovica se zaokružuje na tisuće naniže.
              Zálohe tijekom godine ovise o porezu iz prošle godine. Knjigovođa računa po fakturama i amortizaciji, pa se konačan iznos može razlikovati.
            </div>
          </Card>

          <Card>
            <div style={{ fontWeight: 700, marginBottom: 8 }}>👤 Ako dobit isplatiš sebi (dividenda {Math.round(CZ_DIVIDEND * 100)} %)</div>
            <div style={row}><span>Dobit nakon poreza (procjena za cijelu godinu)</span><b className="num">{Kc(est.afterTax)}</b></div>
            <div style={row}><span>Porez na dividendu (sráží firma)</span><b className="num" style={{ color: S.amber }}>{Kc(est.dividendTax)}</b></div>
            <div style={{ ...row, borderBottom: "none" }}><b>Tebi ostaje na ruke</b><b className="num" style={{ color: S.green }}>{Kc(est.afterTax - est.dividendTax)}</b></div>
            <div style={{ fontSize: 11.5, color: S.sub, marginTop: 4 }}>Dividenda se smije isplatiti tek nakon što valná hromada odobri godišnje izvješće. Porez firma uplaćuje do kraja idućeg mjeseca nakon isplate.</div>
          </Card>

          <Card>
            <div style={{ fontWeight: 700, marginBottom: 8 }}>🧾 DPH {year}. (iz faktura u 🧾 DPH)</div>
            <div style={row}><span>Obračunato za platiti (izdane − primljene fakture)</span><b className="num">{Kc(vatOwed)}</b></div>
            <div style={row}><span>Plaćeno poreznom uredu (s izvoda)</span><b className="num" style={{ color: S.green }}>− {Kc(vatPaid)}</b></div>
            <div style={{ ...row, borderBottom: "none" }}><b>Razlika (otvoreno)</b><b className="num" style={{ color: vatOwed - vatPaid > 0 ? S.amber : S.green }}>{Kc(vatOwed - vatPaid)}</b></div>
            {nextVat && <div style={{ fontSize: 13, marginTop: 4 }}>📅 Sljedeće: <b>{Kc(nextVat.v.payable)}</b> za {monthName(nextVat.mo)} do <b>{fmtDate(nextVat.v.due)}</b></div>}
            <div style={{ fontSize: 11.5, color: S.sub, marginTop: 4 }}>DPH za mjesec se plaća do 25. idućeg, pa je zadnji mjesec obično još otvoren. Primljene fakture s DPH-om upiši u 🧾 DPH da se odbiju.</div>
          </Card>

          <Card>
            <div style={{ fontWeight: 700, marginBottom: 8 }}>📅 Po mjesecima</div>
            <div style={{ ...row, color: S.sub, fontSize: 12, fontWeight: 700 }}><span style={{ flex: 1 }}>Mjesec</span><span style={{ width: 90, textAlign: "right" }}>Prihodi</span><span style={{ width: 90, textAlign: "right" }}>Troškovi</span><span style={{ width: 90, textAlign: "right" }}>Rezultat</span></div>
            {months.map((m, i) => (m.count || invNet[i]) ? (
              <div key={m.month} onClick={() => { setFMonth(m.month); setFCat("all"); setView("transakcije"); }} style={{ ...row, cursor: "pointer" }}>
                <span style={{ flex: 1 }}>{mName(i)}{m.uncat ? <span style={{ color: S.amber }}> · ❔{m.uncat}</span> : null}</span>
                <span className="num" style={{ width: 90, textAlign: "right", color: S.green }}>{Kc(incomeOf(i))}</span>
                <span className="num" style={{ width: 90, textAlign: "right", color: S.red }}>{Kc(m.cost)}</span>
                <b className="num" style={{ width: 90, textAlign: "right" }}>{Kc(incomeOf(i) - m.cost)}</b>
              </div>
            ) : null)}
          </Card>

          <Card>
            <div style={{ fontWeight: 700, marginBottom: 8 }}>🗂 Po kategorijama ({year}.)</div>
            {Object.entries(byCat).sort((a, b) => a[1].gross - b[1].gross).map(([k, v]) => (
              <div key={k} onClick={() => { setFMonth(""); setFCat(k || "none"); setView("transakcije"); }} style={{ ...row, cursor: "pointer" }}>
                <span>{catLabel(k)} <span style={{ color: S.sub, fontSize: 12 }}>({v.count})</span></span>
                <b className="num" style={{ color: v.gross >= 0 ? S.green : S.ink }}>{Kc(v.gross)}</b>
              </div>
            ))}
          </Card>
        </>
      )}

      {view === "obveze" && (() => {
        const pct = (n) => `${String(Math.round(n * 1000) / 10).replace(".", ",")} %`;
        const col = { width: 100, textAlign: "right" };
        const line = (label, now, full, opt = {}) => (
          <div style={{ ...row, ...(opt.bold ? { fontSize: 14.5 } : {}), ...(opt.last ? { borderBottom: "none" } : {}) }}>
            <span style={{ flex: 1 }}>{opt.bold ? <b>{label}</b> : label}{opt.hint && <div style={{ fontSize: 11.5, color: S.sub }}>{opt.hint}</div>}</span>
            <span className="num" style={{ ...col, color: S.sub }}>{now == null ? "—" : Kc(now)}</span>
            <b className="num" style={{ ...col, color: opt.color || S.ink }}>{full == null ? "—" : Kc(full)}</b>
          </div>
        );
        const t = lev.total, y12 = lev.year12;
        const stateNow = t.soc + t.zdr + t.mix + t.dan + est.cit;
        const stateYear = y12.soc + y12.zdr + y12.mix + y12.dan + est.yCit;
        const paidBank = -txs.filter((x) => ["odvodi", "porez_place", "dph", "porez_dobit"].includes(x.category) && x.amount < 0).reduce((s, x) => s + toCzkAmt(x, rate), 0);
        const noLevies = !lev.covered;
        return (
          <>
            {noLevies && (
              <Card style={{ background: S.amberSoft, borderColor: S.amberSoft, fontSize: 13.5 }}>
                Za {year}. još nemam podatke o soc. i zdrav. osiguranju. Učitaj <b>Mzdy PDF</b> s příkazima k úhradě u Obračunu (spremi podsjetnike) ili uvezi izvode i razvrstaj uplate ČSSZ / VZP / poreznom uredu.
              </Card>
            )}
            <Card>
              <div style={{ fontWeight: 700, marginBottom: 2 }}>🧮 Što firma plaća državi za {year}.</div>
              <div style={{ fontSize: 12, color: S.sub, marginBottom: 8 }}>
                {lev.covered ? `Odvodi iz plaća: ${lev.covered} mj. s podacima${lev.factor > 1 ? " → procjena za 12 mjeseci" : ""}.` : ""} Procjena — konačne iznose potvrđuje knjigovođa.
              </div>
              <div style={{ ...row, color: S.sub, fontSize: 12, fontWeight: 700 }}><span style={{ flex: 1 }}></span><span style={col}>do sada</span><span style={col}>cijela godina</span></div>
              {line("🏛 Socijalno osiguranje (ČSSZ)", t.soc, y12.soc, { hint: `firma ${pct(CZ_SOC_FIRM)} + zaposlenik ${pct(CZ_SOC_EMP)} bruto plaće` })}
              {line("🏥 Zdravstveno osiguranje (VZP…)", t.zdr, y12.zdr, { hint: `firma ${pct(CZ_ZDR_FIRM)} + zaposlenik ${pct(CZ_ZDR_EMP)} bruto plaće` })}
              {t.mix > 0 && line("Soc./zdr. — primatelj nije prepoznat", t.mix, y12.mix)}
              {line("🧾 Porez iz plaća (záloha na daň)", t.dan, y12.dan, { hint: "15 % — odbija se radnicima, firma ga uplaćuje" })}
              {line(`🏢 Porez na dobit (${Math.round(CZ_CIT * 100)} %)`, est.cit, est.yCit, { hint: `dobit ${Kc(est.profit)} do sada${est.factor > 1 ? `, ${Kc(est.yProfit)} za cijelu godinu` : ""}` })}
              <div style={{ ...row, fontSize: 15, background: S.amberSoft, borderRadius: 10, padding: "8px 10px", margin: "6px 0", borderBottom: "none" }}>
                <b style={{ flex: 1 }}>Ukupno državi</b>
                <b className="num" style={{ ...col, color: S.sub }}>{Kc(stateNow)}</b>
                <b className="num" style={{ ...col, color: S.amber }}>{Kc(stateYear)}</b>
              </div>
              {line("🧾 DPH (od klijenata, prosljeđuješ državi)", vatOwed, null, { hint: "nije trošak firme — naplaćen je na fakturama", color: S.sub })}
              {line(`👤 Porez na dividendu (${Math.round(CZ_DIVIDEND * 100)} %), ako isplatiš dobit`, null, est.dividendTax, { hint: `od ${Kc(est.afterTax)} dobiti nakon poreza tebi ostaje ${Kc(est.afterTax - est.dividendTax)}`, last: true })}
              {paidBank > 0 && <div style={{ fontSize: 12.5, color: S.sub, marginTop: 6 }}>S izvoda je ove godine već plaćeno (odvodi, porez iz plaća, DPH, zálohe): <b>{Kc(paidBank)}</b></div>}
            </Card>

            <Card>
              <div style={{ fontWeight: 700, marginBottom: 8 }}>👷 Soc. + zdrav. + porez iz plaća — tko to plaća</div>
              <div style={{ ...row, color: S.sub, fontSize: 12, fontWeight: 700 }}><span style={{ flex: 1 }}></span><span style={col}>do sada</span><span style={col}>cijela godina</span></div>
              {line("Trošak firme (dio poslodavca)", lev.split.firm, lev.split12.firm, { hint: "ide preko bruto plaće — to je stvarni dodatni trošak firme", color: S.red })}
              {line("Dio zaposlenika (odbija se od bruto) + porez", lev.split.employees, lev.split12.employees)}
              {line("Ukupno uplaćeno ČSSZ, zdrav. i poreznom uredu", lev.split.all, lev.split12.all, { bold: true })}
              {envelopeSocZdr > 0 && line("Radnicima odbijeno na koverti (soc.-zdr.)", envelopeSocZdr, null, { hint: "zbroj napomena iz Mzdy PDF-a / ručnih soc.-zdr. odbitaka", color: S.green })}
              <div style={{ fontSize: 11.5, color: S.sub, marginTop: 6, marginBottom: 8 }}>
                Podjela firma / zaposlenik je izračunata po zakonskim stopama. Kao vlasnik od dividende ne plaćaš ni socijalno ni zdravstveno, samo 15 % poreza. Ako si i sam na plaći (ili primaš odměnu jednatele), tvoji odvodi su već u ovim brojkama.
              </div>
              <div style={{ ...row, color: S.sub, fontSize: 12, fontWeight: 700 }}><span style={{ flex: 1 }}>Plaće za</span><span style={{ width: 80, textAlign: "right" }}>Soc.</span><span style={{ width: 80, textAlign: "right" }}>Zdr.</span><span style={{ width: 80, textAlign: "right" }}>Porez</span></div>
              {lev.months.map((m, i) => m.source ? (
                <div key={i} style={row}>
                  <span style={{ flex: 1 }}>{mName(i)} <span style={{ fontSize: 11, color: S.sub }}>{m.source === "pdf" ? "Mzdy PDF" : "izvod"}</span></span>
                  <span className="num" style={{ width: 80, textAlign: "right" }}>{Kc(m.soc + (m.mix || 0))}</span>
                  <span className="num" style={{ width: 80, textAlign: "right" }}>{Kc(m.zdr)}</span>
                  <span className="num" style={{ width: 80, textAlign: "right" }}>{Kc(m.dan)}</span>
                </div>
              ) : null)}
            </Card>

            <Card>
              <div style={{ fontWeight: 700, marginBottom: 8 }}>📅 Rokovi i zálohe</div>
              {[
                ["svaki mjesec do 20.", "Soc. i zdrav. osiguranje + porez iz plaća za prošli mjesec (ČSSZ, zdravstvena, porezni ured)"],
                ["svaki mjesec do 25.", "DPH prijava, kontrolní hlášení i uplata DPH za prošli mjesec"],
                ["svaka 3 mj.", "Zákonné pojištění odpovědnosti zaměstnavatele (Kooperativa / ČPP) — mali iznos od bruto plaća"],
                [`ožujak ${year + 1}.`, "Vyúčtování daně ze závislé činnosti (godišnji obračun poreza iz plaća — radi knjigovođa)"],
                [`1. 4. ${year + 1}.`, "Prijava poreza na dobit (DPPO) na papiru — elektronički do 1. 5., preko poreznog savjetnika 1. 7.; do istog roka se plaća razlika poreza"],
                [`uz prijavu`, "Účetní závěrka (bilanca i račun dobiti) u Sbírku listin"],
                ["nakon odluke valne hromade", `Isplata dividende — porez ${Math.round(CZ_DIVIDEND * 100)} % firma uplaćuje do kraja idućeg mjeseca`],
              ].map(([when, what]) => (
                <div key={when} style={{ ...row, justifyContent: "flex-start" }}><b style={{ width: 140, flexShrink: 0 }}>{when}</b><span>{what}</span></div>
              ))}
              <div style={{ marginTop: 10, fontSize: 13.5 }}>
                {adv.count
                  ? <>Ako porez na dobit za {year}. bude oko <b>{Kc(est.yCit)}</b>, nakon prijave plaćaš zálohe <b>{adv.count}× {Kc(adv.amount)}</b> ({adv.dates.join(", ")}) — to je predujam za porez iduće godine.</>
                  : <>Uz procijenjeni porez {Kc(est.yCit)} (do 30 000 Kč) iduće godine nema záloha na porez na dobit.</>}
              </div>
              <div style={{ fontSize: 11.5, color: S.sub, marginTop: 6 }}>Silniční daň se plaća samo za kamione iznad 3,5 t (do 31. 1.). Ako firma ima nekretninu, daň z nemovitých věcí se plaća do 31. 5.</div>
            </Card>
          </>
        );
      })()}

      {view === "transakcije" && (
        <Card>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 8 }}>
            <select value={fMonth} onChange={(e) => setFMonth(e.target.value)} style={sel}>
              <option value="">Cijela godina</option>
              {months.map((m, i) => <option key={m.month} value={m.month}>{mName(i)}</option>)}
            </select>
            <select value={fCat} onChange={(e) => setFCat(e.target.value)} style={sel}>
              <option value="all">Sve kategorije</option>
              <option value="none">❔ Nerazvrstano</option>
              <option value="auto">🤖 Automatski prijedlozi</option>
              {FIN_CATS.map(([k, label]) => <option key={k} value={k}>{label}</option>)}
            </select>
            <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="traži partnera, VS, poruku…" style={{ flex: "1 1 150px", width: "auto", padding: "6px 10px", fontSize: 13 }} />
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 8, fontSize: 12.5 }}>
            <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
              <input type="checkbox" checked={remember} onChange={(e) => setRemember(e.target.checked)} style={{ width: 16, height: 16 }} />
              Zapamti kategoriju za istog partnera
            </label>
            <span style={{ flex: 1 }} />
            {filtered.some((t) => t.category_auto) && <Btn small kind="ghost" onClick={confirmAuto}>✓ Potvrdi prijedloge</Btn>}
            <Btn small kind="ghost" onClick={() => setManual(manual ? null : { date: todayISO(), amount: "", sign: "-", currency: "CZK", counterparty: "", category: "" })}>➕ Ručno</Btn>
          </div>
          {manual && (
            <div style={{ background: S.bg, borderRadius: 10, padding: 10, marginBottom: 8 }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 6 }}>Stavka koja nije na izvodu (npr. gotovina)</div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                <input type="date" value={manual.date} onChange={(e) => setManual({ ...manual, date: e.target.value })} style={{ width: 150 }} />
                <button onClick={() => setManual({ ...manual, sign: manual.sign === "-" ? "+" : "-" })} style={chip(false)}>{manual.sign === "-" ? "− isplata" : "+ uplata"}</button>
                <input inputMode="decimal" value={manual.amount} onChange={(e) => setManual({ ...manual, amount: e.target.value })} placeholder="iznos" style={{ flex: "1 1 90px", width: "auto" }} />
                <CurChips small value={manual.currency} onChange={(c) => setManual({ ...manual, currency: c })} />
                <input value={manual.counterparty} onChange={(e) => setManual({ ...manual, counterparty: e.target.value })} placeholder="opis / partner" style={{ flex: "2 1 140px", width: "auto" }} />
                {catSelect(manual.category, (c) => setManual({ ...manual, category: c }))}
                <Btn small onClick={saveManual}>Spremi</Btn>
              </div>
            </div>
          )}
          <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 4 }}>
            {filtered.length} transakcija · uplate {Kc(filtered.filter((t) => t.amount > 0).reduce((s, t) => s + toCzkAmt(t, rate), 0))} · isplate {Kc(filtered.filter((t) => t.amount < 0).reduce((s, t) => s + toCzkAmt(t, rate), 0))}
          </div>
          {filtered.length === 0 && <Empty text="Nema transakcija za ovaj filter." />}
          {filtered.slice(0, limit).map((t) => {
            const kind = catKind(t.category);
            return (
              <div key={t.id} style={{ padding: "8px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13 }}>
                <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
                  <div style={{ minWidth: 0 }}>
                    <b>{t.counterparty || t.message || t.counter_account || "—"}</b>
                    <div style={{ fontSize: 11.5, color: S.sub, wordBreak: "break-word" }}>
                      {fmtDate(t.tx_date)}{t.counter_account ? " · " + t.counter_account : ""}{t.vs ? " · VS " + t.vs : ""}{t.counterparty && t.message ? " · " + t.message : ""}
                    </div>
                  </div>
                  <b className="num" style={{ color: t.amount > 0 ? S.green : S.ink, whiteSpace: "nowrap" }}>{t.amount > 0 ? "+" : ""}{money(t.amount, t.currency)}</b>
                </div>
                <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap", marginTop: 5 }}>
                  {catSelect(t.category || "", (c) => setCat(t, c),
                    { borderColor: !t.category ? S.amber : t.category_auto ? S.blue : S.line, maxWidth: 230 })}
                  {t.category_auto && <Tag color={S.blue} bg={S.blueSoft}>🤖 prijedlog</Tag>}
                  {(kind === "in" || kind === "cost") && (
                    <span style={{ display: "flex", gap: 3, alignItems: "center", fontSize: 11.5, color: S.sub }}>
                      DPH
                      {[0, 12, 21].map((r) => (
                        <button key={r} onClick={() => patchTx([t.id], { vat_rate: r })} style={{ ...chip(vatOf(t) === r), padding: "2px 7px", fontSize: 11.5 }}>{r}%</button>
                      ))}
                    </span>
                  )}
                  <span style={{ flex: 1 }} />
                  <button onClick={() => delTx(t)} title="Obriši" style={{ background: "none", border: "none", color: S.red, cursor: "pointer", fontSize: 14 }}>✕</button>
                </div>
              </div>
            );
          })}
          {filtered.length > limit && <Btn small kind="ghost" style={{ marginTop: 8 }} onClick={() => setLimit(limit + 300)}>Prikaži još ({filtered.length - limit})</Btn>}
        </Card>
      )}

      {view === "uvoz" && (
        <>
          <Card>
            <div style={{ fontWeight: 700, marginBottom: 6 }}>📥 Učitaj bankovni izvod</div>
            <div style={{ fontSize: 13, color: S.sub, marginBottom: 8 }}>
              U internet bankarstvu izvezi pohyby za mjesec kao <b>CSV</b> (ili Excel). Prepoznaju se Fio, KB, ČSOB, Česká spořitelna, Raiffeisen, Air Bank, Moneta, UniCredit i slični.
              Ako učitaš isti izvod dvaput, iste transakcije se preskaču.
            </div>
            <input type="file" accept=".csv,.txt,.xlsx,.xls" onChange={(e) => { onFile(e.target.files?.[0]); e.target.value = ""; }} />
          </Card>
          {imp && (
            <Card>
              <div style={{ fontWeight: 700, marginBottom: 6 }}>{imp.fileName}</div>
              {imp.notFound && <div style={{ color: S.amber, fontSize: 13, marginBottom: 6 }}>Nisam sam prepoznao zaglavlje — odaberi redak sa zaglavljem i stupce ispod.</div>}
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center", fontSize: 12.5, marginBottom: 6 }}>
                Zaglavlje je u retku
                <select value={imp.header} onChange={(e) => { const h = Number(e.target.value); setImp({ ...imp, header: h, map: guessMapping(imp.rows[h] || []) }); }} style={sel}>
                  {imp.rows.slice(0, 40).map((r, i) => <option key={i} value={i}>{i + 1}: {r.slice(0, 3).join(" | ").slice(0, 40)}</option>)}
                </select>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 6, marginBottom: 8 }}>
                {FIELDS.map(([f, label]) => (
                  <label key={f} style={{ fontSize: 11.5, color: S.sub }}>{label}
                    <select value={imp.map[f] ?? -1} onChange={(e) => setImp({ ...imp, map: { ...imp.map, [f]: Number(e.target.value) } })} style={{ ...sel, width: "100%" }}>
                      <option value={-1}>—</option>
                      {(imp.rows[imp.header] || []).map((h, i) => <option key={i} value={i}>{String(h) || `stupac ${i + 1}`}</option>)}
                    </select>
                  </label>
                ))}
              </div>
              <div style={{ fontSize: 12, color: S.sub, marginBottom: 6 }}>Poruka se slaže iz: {(imp.map.message || []).map((i) => imp.rows[imp.header]?.[i]).join(", ") || "—"}</div>
              <div style={{ fontWeight: 700, fontSize: 13, margin: "6px 0" }}>
                Pronađeno {parsed.txs.length} transakcija{parsed.skipped ? ` (preskočeno ${parsed.skipped} redaka bez datuma/iznosa)` : ""}
              </div>
              {parsed.txs.slice(0, 8).map((t) => (
                <div key={t.hash} style={row}>
                  <span style={{ minWidth: 0 }}>{fmtDate(t.tx_date)} · {t.counterparty || t.message || "—"}<br /><span style={{ fontSize: 11.5, color: S.sub }}>{catLabel(suggestCategory(t, rules))}</span></span>
                  <b className="num" style={{ color: t.amount > 0 ? S.green : S.ink, whiteSpace: "nowrap" }}>{money(t.amount, t.currency)}</b>
                </div>
              ))}
              <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                <Btn small onClick={doImport} disabled={busy || !parsed.txs.length}>Uvezi {parsed.txs.length} transakcija</Btn>
                <Btn small kind="ghost" onClick={() => setImp(null)}>Odustani</Btn>
              </div>
            </Card>
          )}
          <Card>
            <div style={{ fontWeight: 700, marginBottom: 6 }}>Prethodni uvozi</div>
            {imports.length === 0 ? <div style={{ color: S.sub, fontSize: 13 }}>Još ništa nije uvezeno.</div> : imports.map((im) => (
              <div key={im.id} style={row}>
                <span>{im.file_name}<br /><span style={{ fontSize: 11.5, color: S.sub }}>{fmtDate(im.date_from)} – {fmtDate(im.date_to)} · {im.row_count} transakcija · uvezeno {fmtDate(String(im.created_at).slice(0, 10))}</span></span>
                <button onClick={() => delImport(im)} style={{ background: "none", border: "none", color: S.red, cursor: "pointer", fontSize: 15 }}>✕</button>
              </div>
            ))}
          </Card>
        </>
      )}

      {view === "pravila" && (
        <Card>
          <div style={{ fontWeight: 700, marginBottom: 6 }}>⚙️ Pravila za automatsko razvrstavanje</div>
          <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 8 }}>
            Ako ime partnera, broj računa ili poruka sadrži zadani tekst, transakcija dobije tu kategoriju. Pravila nastaju sama kad promijeniš kategoriju (uz kvačicu „Zapamti“).
            Uz tvoja pravila postoje i ugrađena (ČSSZ, VZP, finanční úřad, poplatek, benzinske…).
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 8 }}>
            <input value={ruleForm.pattern} onChange={(e) => setRuleForm({ ...ruleForm, pattern: e.target.value })} placeholder="tekst, npr. Hotel Praha ili 123456789/0100" style={{ flex: "2 1 160px", width: "auto" }} />
            {catSelect(ruleForm.category, (c) => setRuleForm({ ...ruleForm, category: c }))}
            <Btn small onClick={addRule} disabled={!ruleForm.pattern.trim() || !ruleForm.category}>Dodaj</Btn>
          </div>
          {rules.map((r) => (
            <div key={r.id} style={row}>
              <span>„{r.pattern}“ → {catLabel(r.category)}</span>
              <button onClick={() => delRule(r)} style={{ background: "none", border: "none", color: S.red, cursor: "pointer", fontSize: 15 }}>✕</button>
            </div>
          ))}
          <Btn small kind="ghost" style={{ marginTop: 10 }} onClick={applyRules}>🤖 Primijeni pravila na nerazvrstane i prijedloge ({year}.)</Btn>
        </Card>
      )}
    </div>
  );
}
function AdminPanels({ data, api, panel, setPanel, onOpenWorker, onOpenProfile }) {
  const [firm, setFirm] = useState({ company_name: "", address: "", oib: "", iban: "", czk_rate: "25", czk_rate_auto: true, weekend_pct: "0", holiday_pct: "0", expiry_warn_days: "30", payday_day: "5", payday_reminder_days: "3" });
  const [auditQ, setAuditQ] = useState("");
  const [commEdit, setCommEdit] = useState({});
  const [objCommEdit, setObjCommEdit] = useState({}); // "profileId:objectId" -> { rate, cur }
  const [apprSel, setApprSel] = useState(() => new Set());
  const [apprBusy, setApprBusy] = useState(false);
  const [vatMonth, setVatMonth] = useState(prevMonthKey(curMonth()));
  const [piForm, setPiForm] = useState({ supplier: "", number: "", duzp: todayISO(), base: "", vat: "", currency: "CZK" });
  const [lastBackup, setLastBackup] = useState(() => { try { return localStorage.getItem("evidencija_last_backup"); } catch { return null; } });
  const daysSinceBackup = lastBackup ? Math.floor((Date.now() - new Date(lastBackup).getTime()) / 86400000) : null;
  const doBackup = () => {
    exportAllData(data);
    const now = new Date().toISOString();
    try { localStorage.setItem("evidencija_last_backup", now); } catch {}
    setLastBackup(now);
  };
  const wName = (id) => data.workers.find((w) => w.id === id)?.name || "Obrisan radnik";
  const pendingPayouts = [...(data.payouts || [])].filter((p) => !p.approved).sort((a, b) => a.month.localeCompare(b.month) || wName(a.workerId).localeCompare(wName(b.workerId), "hr"));
  const liveNetFor = (workerId, mo) => {
    const row = calcRows(data, (d) => monthKey(d) === mo, new Set()).find((r) => r.w.id === workerId);
    return row ? { net: row.net, netKc: row.czk.net } : { net: 0, netKc: 0 };
  };
  const allApprSelected = pendingPayouts.length > 0 && pendingPayouts.every((p) => apprSel.has(p.id));
  const toggleApprSelectAll = () => setApprSel(allApprSelected ? new Set() : new Set(pendingPayouts.map((p) => p.id)));
  const toggleApprSel = (id) => setApprSel((prev) => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const approveList = async (list) => {
    setApprBusy(true);
    for (const p of list) await api.approvePayout(p, wName(p.workerId));
    setApprSel(new Set());
    setApprBusy(false);
  };
  useEffect(() => setFirm({
    company_name: data.settings.company_name || "", address: data.settings.address || "",
    oib: data.settings.oib || "", iban: data.settings.iban || "", czk_rate: String(data.settings.czk_rate || 25), czk_rate_auto: data.settings.czk_rate_auto !== false,
    weekend_pct: String(data.settings.weekend_pct || 0), holiday_pct: String(data.settings.holiday_pct || 0),
    expiry_warn_days: String(data.settings.expiry_warn_days || 30),
    payday_day: String(data.settings.payday_day || 5), payday_reminder_days: String(data.settings.payday_reminder_days || 3),
  }), [data.settings]);

  return (
    <div style={{ marginBottom: 12 }}>
      {panel && (
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
          <Btn small kind="ghost" onClick={() => setPanel("")}>{tr("← Zatvori")}</Btn>
          <div style={{ fontWeight: 800, fontSize: 15 }}>{(ADMIN_MENU.find((m) => m[0] === panel) || [])[2]}</div>
        </div>
      )}

      {panel === "approvals" && (
        <Card style={{ marginTop: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("⏳ Isplate koje čekaju odobrenje — svi mjeseci i radnici na jednom mjestu")}</div>
          {pendingPayouts.length === 0 ? (
            <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Nema ničega na čekanju. 🎉")}</div>
          ) : (
            <>
              <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, cursor: "pointer", marginBottom: 8 }}>
                <input type="checkbox" checked={allApprSelected} onChange={toggleApprSelectAll} style={{ width: 18, height: 18 }} />
                {tr("Odaberi sve (")}{pendingPayouts.length})
              </label>
              {pendingPayouts.map((p) => {
                const live = liveNetFor(p.workerId, p.month);
                const mismatch = round2(p.amount || 0) !== round2(live.net || 0) || round2(p.amountKc || 0) !== round2(live.netKc || 0);
                return (
                <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13.5, flexWrap: "wrap" }}>
                  <input type="checkbox" checked={apprSel.has(p.id)} onChange={() => toggleApprSel(p.id)} style={{ width: 17, height: 17 }} />
                  <div style={{ flex: 1 }}>
                    <span onClick={() => onOpenWorker && onOpenWorker(p.workerId)} style={{ fontWeight: 700, cursor: onOpenWorker ? "pointer" : "default", color: onOpenWorker ? S.blue : "inherit" }}>
                      {wName(p.workerId)}
                    </span>
                    <span className="num" style={{ color: S.sub }}> · {p.month} {tr("· predloženo")}{" "}{fmtDate(p.paidAt)}</span>
                  </div>
                  <span className="num" style={{ fontWeight: 700, color: S.amber }}>
                    {[p.amount ? eur(p.amount) : "", p.amountKc ? czk(p.amountKc) : ""].filter(Boolean).join(" + ") || eur(0)}
                  </span>
                  <Btn small onClick={() => approveList([p])} disabled={apprBusy}>{tr("✓ Odobri")}</Btn>
                  <button onClick={() => api.unmarkPaid(p, wName(p.workerId))} title={tr("Odbaci")} style={{ background: "none", border: "none", color: S.red, fontSize: 16, cursor: "pointer", padding: 4 }}>✕</button>
                  {mismatch && (
                    <div style={{ width: "100%", display: "flex", alignItems: "center", gap: 8, background: S.redSoft, borderRadius: 8, padding: "5px 9px" }}>
                      <span style={{ fontSize: 11.5, color: S.red, fontWeight: 700, flex: 1 }}>
                        {tr("⚠️ Promijenjeno nakon prijedloga — sad bi bilo")}{" "}{[live.net ? eur(live.net) : "", live.netKc ? czk(live.netKc) : ""].filter(Boolean).join(" + ") || eur(0)}
                      </span>
                      <Btn small onClick={() => api.refreshPayout(p, live.net, live.netKc, wName(p.workerId))}>{tr("🔄 Osvježi")}</Btn>
                    </div>
                  )}
                </div>
                );
              })}
              <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                {apprSel.size > 0 && (
                  <Btn small onClick={() => approveList(pendingPayouts.filter((p) => apprSel.has(p.id)))} disabled={apprBusy}>
                    {apprBusy ? tr("Odobravam…") : tr("✓ Odobri odabrane ({0})", apprSel.size)}
                  </Btn>
                )}
                <Btn small kind={apprSel.size > 0 ? "ghost" : "primary"} onClick={() => approveList(pendingPayouts)} disabled={apprBusy}>
                  {apprBusy ? tr("Odobravam…") : tr("✓✓ Odobri sve ({0})", pendingPayouts.length)}
                </Btn>
              </div>
            </>
          )}
        </Card>
      )}

      {panel === "unpaid" && <UnpaidPanel data={data} api={api} onOpenWorker={onOpenWorker} />}
      {panel === "candidates" && <CandidatesPanel data={data} api={api} />}
      {panel === "finance" && <FinancePanel data={data} api={api} />}
      {panel === "cash" && <CashPanel data={data} api={api} />}

      {panel === "users" && (
        <Card style={{ marginTop: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("👥 Zaposlenici s pristupom")}</div>
          {[...data.profiles].sort((a, b) => (a.name || "").localeCompare(b.name || "", "hr")).map((p) => {
            const edit = commEdit[p.id];
            const cVal = edit !== undefined ? edit.rate : String(p.hourly_commission || "");
            const cCur = edit !== undefined ? edit.cur : (p.commission_currency || "EUR");
            const myObjIds = [...new Set((data.objectMembers || []).filter((m) => m.user_id === p.id).map((m) => m.object_id))];
            const myObjects = myObjIds.map((oid) => data.objects.find((o) => o.id === oid)).filter(Boolean).sort((a, b) => a.name.localeCompare(b.name, "hr"));
            const mk = curMonth();
            const hoursForObj = (oid) => round2(data.logs.filter((l) => l.objectId === oid && monthKey(l.date) === mk).reduce((s, l) => s + l.hours, 0));
            const rateForObj = (oid) => {
              const ov = (data.commissionRates || []).find((c) => c.userId === p.id && c.objectId === oid);
              return ov ? { rate: ov.rate, cur: ov.currency } : { rate: parseNum(cVal) || 0, cur: cCur };
            };
            const totalsByCur = {};
            myObjects.forEach((o) => {
              const { rate, cur } = rateForObj(o.id);
              totalsByCur[cur] = round2((totalsByCur[cur] || 0) + hoursForObj(o.id) * rate);
            });
            return (
              <div key={p.id} style={{ padding: "9px 0", borderBottom: `1px solid ${S.line}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span onClick={() => onOpenProfile && onOpenProfile(p.id)} style={{ flex: 1, fontWeight: 600, fontSize: 14, cursor: onOpenProfile ? "pointer" : "default", color: onOpenProfile ? S.blue : "inherit" }}>
                    {p.role === "admin" ? "👑 " : ""}{p.name || tr("(bez imena)")} {onOpenProfile && <span style={{ fontWeight: 700 }}>›</span>}
                  </span>
                  <select value={p.role} onChange={(e) => api.setRole(p, e.target.value)} style={{ width: "auto", padding: "6px 8px", fontSize: 13 }}>
                    <option value="employee">{tr("Zaposlenik")}</option>
                    <option value="admin">{tr("Admin")}</option>
                  </select>
                </div>
                {p.role !== "admin" && (
                  <div style={{ marginTop: 6 }}>
                    <div style={{ fontSize: 11.5, color: S.sub, marginBottom: 3 }}>{tr("Zadana provizija po satu (za objekte bez posebne cijene ispod)")}</div>
                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                      <input inputMode="decimal" value={cVal} onChange={(e) => setCommEdit((p2) => ({ ...p2, [p.id]: { rate: e.target.value, cur: cCur } }))}
                        placeholder={tr("npr. 0.50")} style={{ flex: 1 }} />
                      <CurChips small value={cCur} onChange={(v) => setCommEdit((p2) => ({ ...p2, [p.id]: { rate: cVal, cur: v } }))} />
                      <Btn small onClick={() => api.setCommission(p, round2(parseNum(cVal) || 0), cCur)}>{tr("Spremi")}</Btn>
                    </div>

                    {myObjects.length > 0 && (
                      <div style={{ marginTop: 8, borderTop: `1px dashed ${S.line}`, paddingTop: 6 }}>
                        <div style={{ fontSize: 11.5, color: S.sub, marginBottom: 4 }}>{tr("Posebna cijena po objektu (opcionalno — nadjačava zadanu gore)")}</div>
                        {myObjects.map((o) => {
                          const ov = (data.commissionRates || []).find((c) => c.userId === p.id && c.objectId === o.id);
                          const key = `${p.id}:${o.id}`;
                          const oedit = objCommEdit[key];
                          const oVal = oedit !== undefined ? oedit.rate : (ov ? String(ov.rate) : "");
                          const oCur = oedit !== undefined ? oedit.cur : (ov ? ov.currency : cCur);
                          const h = hoursForObj(o.id);
                          const { rate: effRate, cur: effCur } = rateForObj(o.id);
                          return (
                            <div key={o.id} style={{ display: "flex", alignItems: "center", gap: 6, padding: "4px 0", fontSize: 12.5, flexWrap: "wrap" }}>
                              <span style={{ flex: "1 1 100px" }}>{o.name} <span className="num" style={{ color: S.sub }}>({fmtH(h)})</span></span>
                              <input inputMode="decimal" value={oVal} onChange={(e) => setObjCommEdit((p2) => ({ ...p2, [key]: { rate: e.target.value, cur: oCur } }))}
                                placeholder={cVal || "0"} style={{ width: 64, padding: "5px 6px" }} />
                              <CurChips small value={oCur} onChange={(v) => setObjCommEdit((p2) => ({ ...p2, [key]: { rate: oVal, cur: v } }))} />
                              <Btn small kind="ghost" onClick={() => api.setCommissionRate(o, p, round2(parseNum(oVal) || 0), oCur)}>✓</Btn>
                              {ov && <button onClick={() => api.delCommissionRate(o, p)} title={tr("Obriši (vrati na zadanu)")} style={{ background: "none", border: "none", color: S.red, fontSize: 15, cursor: "pointer", padding: 2 }}>✕</button>}
                              <span className="num" style={{ fontWeight: 700, color: S.green, minWidth: 62, textAlign: "right" }}>{money(round2(h * effRate), effCur)}</span>
                            </div>
                          );
                        })}
                      </div>
                    )}

                    {Object.values(totalsByCur).some((v) => v !== 0) && (
                      <div style={{ fontSize: 12.5, color: S.sub, marginTop: 6, fontWeight: 700 }}>
                        {tr("Ovaj mjesec ukupno za isplatu:")}{" "}{Object.entries(totalsByCur).filter(([, v]) => v !== 0).map(([c, v]) => money(v, c)).join(" + ")}
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
          <AddEmployee data={data} api={api} />
        </Card>
      )}

      {panel === "vat" && (() => {
        const [vy, vm] = vatMonth.split("-").map(Number);
        const prevVM = () => setVatMonth(vm === 1 ? `${vy - 1}-12` : `${vy}-${String(vm - 1).padStart(2, "0")}`);
        const nextVM = () => setVatMonth(vm === 12 ? `${vy + 1}-01` : `${vy}-${String(vm + 1).padStart(2, "0")}`);
        const objName = (id) => data.objects.find((o) => o.id === id)?.name || "Obrisan objekt";
        const v = vatSummary(data, vatMonth);
        const dd = daysUntil(v.due);
        const outSorted = [...v.out].sort((a, b) => (a.duzp || a.issueDate).localeCompare(b.duzp || b.issueDate));
        const fillVat21 = () => { const base = parseNum(piForm.base); if (base > 0) setPiForm({ ...piForm, vat: String(round2(base * 0.21)) }); };
        const submitPi = async () => {
          const base = round2(parseNum(piForm.base) || 0), vat = round2(parseNum(piForm.vat) || 0);
          if (!piForm.supplier.trim() || !piForm.duzp || !(vat > 0)) return;
          const ok = await api.addPurchaseInvoice({ ...piForm, supplier: piForm.supplier.trim(), number: piForm.number.trim(), base, vat });
          if (ok) setPiForm({ supplier: "", number: "", duzp: piForm.duzp, base: "", vat: "", currency: piForm.currency });
        };
        const row = { display: "flex", justifyContent: "space-between", gap: 8, padding: "5px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13 };
        return (
          <Card style={{ marginTop: 8 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div style={{ fontWeight: 700 }}>{tr("🧾 DPH za razdoblje")}</div>
              <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                <button onClick={prevVM} style={{ background: "#fff", border: `1px solid ${S.line}`, borderRadius: 8, width: 26, height: 26, cursor: "pointer", fontWeight: 700 }}>‹</button>
                <div style={{ fontWeight: 700, padding: "0 4px", fontSize: 13.5 }}>{mName(vm - 1)} {vy}.</div>
                <button onClick={nextVM} style={{ background: "#fff", border: `1px solid ${S.line}`, borderRadius: 8, width: 26, height: 26, cursor: "pointer", fontWeight: 700 }}>›</button>
              </div>
            </div>

            <div style={{ background: v.payable > 0 ? S.amberSoft : S.greenSoft, borderRadius: 12, padding: "10px 12px", marginBottom: 10 }}>
              <div style={{ ...row, borderBottom: "none", padding: "2px 0" }}><span>{tr("DPH na izdanim fakturama")}</span><b className="num">{czk(v.outVat)}</b></div>
              <div style={{ ...row, borderBottom: "none", padding: "2px 0" }}><span>{tr("− odbitak s primljenih faktura")}</span><b className="num">{czk(v.inVat)}</b></div>
              <div style={{ ...row, borderBottom: "none", borderTop: `1px dashed ${S.line}`, paddingTop: 6, marginTop: 4, fontSize: 15 }}>
                <b>{v.payable >= 0 ? tr("Za platiti") : tr("Nadměrný odpočet (vraća se vama)")}</b>
                <b className="num" style={{ color: v.payable > 0 ? S.amber : S.green, fontSize: 16 }}>{czk(Math.abs(v.payable))}</b>
              </div>
              <div style={{ fontSize: 13, marginTop: 4, color: dd < 0 ? S.red : S.ink }}>
                {tr("📅 Rok za prijavu i plaćanje:")}{" "}<b>{fmtDate(v.due)}</b>{" "}
                <span style={{ color: dd < 0 ? S.red : S.sub }}>({dd === 0 ? tr("danas") : dd > 0 ? tr("za {0} {1}", dd, dd === 1 ? "dan" : "dana") : tr("prošao prije {0} {1}", -dd, -dd === 1 ? "dan" : "dana")})</span>
              </div>
              {v.eurUsed && <div style={{ fontSize: 11.5, color: S.amber, marginTop: 4 }}>{tr("Dio DPH-a je u €, preračunato po tečaju")}{" "}{v.rate} {tr("— za prijavu vrijedi tečaj ČNB na dan DUZP.")}</div>}
            </div>
            <div style={{ fontSize: 11.5, color: S.sub, marginBottom: 10 }}>
              {tr("Faktura ide u mjesec prema DUZP-u. Ako DUZP nije upisan, prema razdoblju fakture (\"za měsíc …\"), a tek onda prema datumu izdavanja. Rok vrijedi za mjesečnog platitelja DPH (25. idućeg mjeseca, pomaknuto na radni dan). Ovo je pomoć za planiranje, konačan iznos potvrdi s knjigovođom.")}
            </div>

            <div style={{ fontWeight: 700, fontSize: 13.5, margin: "4px 0" }}>{tr("Izdane fakture (")}{v.out.length})</div>
            {outSorted.length === 0 ? (
              <div style={{ color: S.sub, fontSize: 13 }}>{tr("Nema izdanih faktura za ovo razdoblje. Uvezi PDF-ove iz Vyfakturuja u kartici Objekti → \"📥 Uvoz faktura\".")}</div>
            ) : outSorted.map((i) => (
              <div key={i.id} style={row}>
                <span>{objName(i.objectId)} · {i.number}<br /><span style={{ fontSize: 11.5, color: S.sub }}>{i.duzp ? "DUZP " + fmtDate(i.duzp) : "izdano " + fmtDate(i.issueDate)} {tr("· po")}{" "}{invoiceTaxBasis(i)}</span></span>
                <span className="num" style={{ textAlign: "right" }}>{money(i.amount, i.currency)}<br /><b style={{ color: S.amber }}>{tr("DPH")}{" "}{money(i.vatAmount || 0, i.vatCurrency)}</b></span>
              </div>
            ))}

            <div style={{ fontWeight: 700, fontSize: 13.5, margin: "12px 0 4px" }}>{tr("Primljene fakture s DPH-om — odbitak (")}{v.inn.length})</div>
            {v.inn.map((pi) => (
              <div key={pi.id} style={row}>
                <span>{pi.supplier}{pi.number ? " · " + pi.number : ""}<br /><span style={{ fontSize: 11.5, color: S.sub }}>{tr("DUZP")}{" "}{fmtDate(pi.duzp)}{pi.base ? " · základ " + money(pi.base, pi.currency) : ""}</span></span>
                <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <b className="num" style={{ color: S.green }}>{tr("DPH")}{" "}{money(pi.vat, pi.currency)}</b>
                  <button onClick={() => window.confirm(tr("Obrisati fakturu {0} {1}?", pi.supplier, pi.number)) && api.delPurchaseInvoice(pi)} style={{ background: "none", border: "none", color: S.red, cursor: "pointer", fontSize: 15 }}>✕</button>
                </span>
              </div>
            ))}
            <div style={{ background: S.bg, borderRadius: 10, padding: 10, marginTop: 8 }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 6 }}>{tr("➕ Dodaj primljenu fakturu (gorivo, najam, knjigovođa, oprema…)")}</div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
                <input value={piForm.supplier} onChange={(e) => setPiForm({ ...piForm, supplier: e.target.value })} placeholder={tr("dobavljač")} style={{ flex: "2 1 140px" }} />
                <input value={piForm.number} onChange={(e) => setPiForm({ ...piForm, number: e.target.value })} placeholder={tr("broj fakture")} style={{ flex: "1 1 100px" }} />
                <label style={{ flex: "1 1 130px", fontSize: 11, color: S.sub }}>{tr("DUZP")}
                  <input type="date" value={piForm.duzp} onChange={(e) => setPiForm({ ...piForm, duzp: e.target.value })} />
                </label>
              </div>
              <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                <input inputMode="decimal" value={piForm.base} onChange={(e) => setPiForm({ ...piForm, base: e.target.value })} placeholder={tr("základ bez DPH")} style={{ flex: "1 1 100px" }} />
                <Btn small kind="ghost" onClick={fillVat21}>21 %</Btn>
                <input inputMode="decimal" value={piForm.vat} onChange={(e) => setPiForm({ ...piForm, vat: e.target.value })} placeholder={tr("DPH")} style={{ flex: "1 1 90px" }} />
                <CurChips small value={piForm.currency} onChange={(c) => setPiForm({ ...piForm, currency: c })} />
                <Btn small onClick={submitPi} disabled={!piForm.supplier.trim() || !(parseNum(piForm.vat) > 0)}>{tr("Spremi")}</Btn>
              </div>
            </div>
          </Card>
        );
      })()}

      {panel === "audit" && (
        <Card style={{ marginTop: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("📜 Tko je što radio")}</div>
          {data.audit.length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Još nema zabilježenih radnji.")}</div> : (
            <>
              <input value={auditQ} onChange={(e) => setAuditQ(e.target.value)} placeholder={tr("🔍 Traži po osobi ili radnji…")} style={{ marginBottom: 8 }} />
              {(() => {
                const t = auditQ.trim().toLowerCase();
                const filtered = t ? data.audit.filter((a) => (a.user_name || "").toLowerCase().includes(t) || (a.action || "").toLowerCase().includes(t)) : data.audit;
                return filtered.length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Nema rezultata (pretražuje se zadnjih")}{" "}{data.audit.length} {tr("zabilježenih radnji).")}</div>
                  : filtered.map((a) => (
                    <div key={a.id} style={{ padding: "6px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13 }}>
                      <span className="num" style={{ color: S.sub }}>{fmtDT(a.at)}</span>{" · "}<b>{a.user_name}</b> — {a.action}
                    </div>
                  ));
              })()}
            </>
          )}
        </Card>
      )}

      {panel === "logins" && (
        <Card style={{ marginTop: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("🔐 Povijest prijava")}</div>
          {(data.loginLog || []).length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Još nema zabilježenih prijava.")}</div>
            : (data.loginLog || []).map((l) => (
              <div key={l.id} style={{ padding: "6px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13 }}>
                <span className="num" style={{ color: S.sub }}>{fmtDT(l.at)}</span>{" · "}<b>{l.userName || tr("(bez imena)")}</b>
                <div style={{ fontSize: 11, color: S.sub }}>{l.userAgent}</div>
              </div>
            ))}
        </Card>
      )}

      {panel === "errors" && (
        <Card style={{ marginTop: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("🐛 Greške aplikacije")}</div>
          {(data.errorLog || []).length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Nema zabilježenih grešaka. 🎉")}</div>
            : (data.errorLog || []).map((e) => (
              <div key={e.id} style={{ padding: "6px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13 }}>
                <span className="num" style={{ color: S.sub }}>{fmtDT(e.at)}</span>{" · "}<b>{e.userName || tr("(bez imena)")}</b>
                <div style={{ color: S.red, fontWeight: 600 }}>{e.message}</div>
                {e.url && <div style={{ fontSize: 11, color: S.sub }}>{e.url}</div>}
              </div>
            ))}
        </Card>
      )}

      {panel === "trash" && (
        <Card style={{ marginTop: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("🗑 Obrisano — može se vratiti")}</div>
          {data.trash.length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Koš je prazan.")}</div>
            : data.trash.map((t) => (
              <div key={t.table + t.row.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderBottom: `1px solid ${S.line}` }}>
                <div style={{ flex: 1, fontSize: 13 }}>
                  <div style={{ fontWeight: 600 }}>{t.label}</div>
                  <div className="num" style={{ fontSize: 11.5, color: S.sub }}>{fmtDT(t.row.deleted_at)}</div>
                </div>
                <Btn small kind="ghost" onClick={() => api.restore(t)} style={{ color: S.green, fontWeight: 700 }}>{tr("↩ Vrati")}</Btn>
              </div>
            ))}
        </Card>
      )}

      {panel === "firm" && (
        <Card style={{ marginTop: 8 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("⚙️ Podaci firme (za specifikacije i obračune)")}</div>
          <Field label={tr("Naziv firme")}><input value={firm.company_name} onChange={(e) => setFirm({ ...firm, company_name: e.target.value })} placeholder={tr("npr. Moja Firma d.o.o.")} /></Field>
          <Field label={tr("Adresa")}><input value={firm.address} onChange={(e) => setFirm({ ...firm, address: e.target.value })} /></Field>
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}><Field label={tr("IČO (CZ) ili OIB (HR)")}><input value={firm.oib} onChange={(e) => setFirm({ ...firm, oib: e.target.value })} /></Field></div>
            <div style={{ flex: 1.4 }}><Field label={tr("IBAN")}><input value={firm.iban} onChange={(e) => setFirm({ ...firm, iban: e.target.value })} /></Field></div>
          </div>
          <Field label={tr("Tečaj 1 € = ? Kč (za izračun iz Mzdy PDF-a i zbrojeve u €)")}>
            <input inputMode="decimal" value={firm.czk_rate} onChange={(e) => setFirm({ ...firm, czk_rate: e.target.value, czk_rate_auto: false })} placeholder={tr("npr. 25")} />
          </Field>
          <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, marginTop: -4, marginBottom: 10, cursor: "pointer" }}>
            <input type="checkbox" checked={!!firm.czk_rate_auto} onChange={(e) => setFirm({ ...firm, czk_rate_auto: e.target.checked })} style={{ width: 17, height: 17 }} />
            <span>{tr("Automatski službeni tečaj ČNB (radnim danima popodne)")}
              {data.settings.czk_rate_date && <span style={{ color: S.sub }}> {tr("· zadnji:")}{" "}{fmtDate(data.settings.czk_rate_date)}</span>}
            </span>
          </label>
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}>
              <Field label={tr("Dodatak vikendom (%)")}>
                <input inputMode="decimal" value={firm.weekend_pct} onChange={(e) => setFirm({ ...firm, weekend_pct: e.target.value })} placeholder={tr("npr. 25")} />
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label={tr("Dodatak praznikom (%)")}>
                <input inputMode="decimal" value={firm.holiday_pct} onChange={(e) => setFirm({ ...firm, holiday_pct: e.target.value })} placeholder={tr("npr. 50")} />
              </Field>
            </div>
          </div>
          <div style={{ fontSize: 12, color: S.sub, margin: "-6px 0 10px" }}>
            {tr("Automatski se dodaje na zaradu radnika za sate upisane vikendom/na praznik (0 = bez dodatka). Ne utječe na naplatu objekta.")}
          </div>
          <Field label={tr("Upozori na istek dokumenata koliko dana unaprijed")}>
            <input inputMode="decimal" value={firm.expiry_warn_days} onChange={(e) => setFirm({ ...firm, expiry_warn_days: e.target.value })} placeholder={tr("npr. 30")} />
          </Field>
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}>
              <Field label={tr("Dan isplate u mjesecu")}>
                <input inputMode="decimal" value={firm.payday_day} onChange={(e) => setFirm({ ...firm, payday_day: e.target.value })} placeholder={tr("npr. 5")} />
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label={tr("Podsjeti koliko dana prije")}>
                <input inputMode="decimal" value={firm.payday_reminder_days} onChange={(e) => setFirm({ ...firm, payday_reminder_days: e.target.value })} placeholder={tr("npr. 3")} />
              </Field>
            </div>
          </div>
          <div style={{ fontSize: 12, color: S.sub, margin: "-6px 0 10px" }}>
            {tr("Ako nekome ostane neodobrena isplata, dobit ćeš obavijest nekoliko dana prije dana isplate.")}
          </div>
          <Btn small onClick={() => api.saveSettings({ ...firm, czk_rate: parseNum(firm.czk_rate) || 25, czk_rate_auto: !!firm.czk_rate_auto, weekend_pct: parseNum(firm.weekend_pct) || 0, holiday_pct: parseNum(firm.holiday_pct) || 0, expiry_warn_days: parseNum(firm.expiry_warn_days) || 30, payday_day: parseNum(firm.payday_day) || 5, payday_reminder_days: parseNum(firm.payday_reminder_days) || 3 })}>{tr("Spremi")}</Btn>
          <div style={{ borderTop: `1px dashed ${S.line}`, marginTop: 14, paddingTop: 12 }}>
            <div style={{ fontWeight: 700, marginBottom: 6 }}>{tr("⬇ Backup")}</div>
            <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 8 }}>{tr("Sve podatke (radnici, objekti, svi upisani sati, sve isplate, isplaćeni mjeseci) u jednu Excel datoteku, neovisno o odabranom mjesecu.")}</div>
            {(daysSinceBackup === null || daysSinceBackup >= 30) ? (
              <div style={{ fontSize: 12.5, color: S.amber, fontWeight: 600, marginBottom: 8 }}>
                ⚠️ {daysSinceBackup === null ? tr("Još nikad nisi izvezla backup na ovom uređaju.") : tr("Zadnji backup prije {0} dana.", daysSinceBackup)} {tr("Preporuka: barem jednom mjesečno.")}
              </div>
            ) : (
              <div style={{ fontSize: 12, color: S.sub, marginBottom: 8 }}>{tr("Zadnji backup:")}{" "}{daysSinceBackup === 0 ? tr("danas") : tr("prije {0} dana", daysSinceBackup)}.</div>
            )}
            <Btn small kind="excel" onClick={doBackup}>{tr("📊 Izvezi sve podatke (Excel)")}</Btn>
            <div style={{ fontSize: 11, color: S.sub, marginTop: 6 }}>{tr("Ovo je podsjetnik na ovom uređaju/pregledniku — pravi automatski backup na Google Drive/e-mail treba dodatnu poslužiteljsku postavku (mogu pripremiti kod, ali postavljanje ide preko tvog Supabase/Google naloga).")}</div>
          </div>
        </Card>
      )}
    </div>
  );
}

/* ================================================================== */
/*  RADNICI + OBJEKTI + upozorenja na dokumente                        */
/* ================================================================== */
function expiryWarnings(workers, days) {
  const soon = addDaysISO(todayISO(), days || 30);
  const out = [];
  workers.filter((w) => !w.archived).forEach((w) => {
    [["permitExpiry", "radna dozvola"], ["contractExpiry", "ugovor"]].forEach(([k, label]) => {
      const d = w[k];
      if (d && d <= soon) out.push({ w, label, date: d, past: d < todayISO() });
    });
  });
  return out.sort((a, b) => a.date.localeCompare(b.date));
}

const RATE_STALE_DAYS = 365;
// samo za radnike koji IMAJU zabilježenu promjenu satnice — ako je nikad nisu imali, ne znamo otkad stoji ista pa ne upozoravamo
function rateStaleWarnings(workers, rateChanges) {
  const cutoff = addDaysISO(todayISO(), -RATE_STALE_DAYS);
  const out = [];
  workers.filter((w) => !w.archived).forEach((w) => {
    const changes = (rateChanges || []).filter((r) => r.workerId === w.id).sort((a, b) => b.from.localeCompare(a.from));
    if (!changes.length) return;
    if (changes[0].from <= cutoff) out.push({ w, since: changes[0].from });
  });
  return out.sort((a, b) => a.since.localeCompare(b.since));
}

/* ================================================================== */
/*  📌 DANAS — sve što traži pažnju na jednom mjestu                   */
/* ================================================================== */
function TodaySection({ icon, title, count, tone = "amber", children, items, render, limit = 5, action }) {
  const [all, setAll] = useState(false);
  const col = { red: [S.red, S.redSoft], amber: [S.amber, S.amberSoft], blue: [S.blue, S.blueSoft], green: [S.green, S.greenSoft] }[tone];
  const shown = items ? (all ? items : items.slice(0, limit)) : null;
  return (
    <Card style={{ padding: "12px 14px" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
        <span style={{ fontSize: 18 }}>{icon}</span>
        <span style={{ fontWeight: 700, flex: 1, fontSize: 14.5 }}>{title}</span>
        {count !== undefined && <Tag color={col[0]} bg={col[1]}>{count}</Tag>}
      </div>
      {children}
      {shown && shown.map(render)}
      {items && items.length > limit && (
        <button onClick={() => setAll(!all)} style={{ background: "none", border: "none", color: S.blue, fontWeight: 700, fontSize: 13, cursor: "pointer", padding: "6px 0 0" }}>
          {all ? tr("Prikaži manje") : tr("Prikaži sve ({0})", items.length)}
        </button>
      )}
      {action && <div style={{ marginTop: 8 }}>{action}</div>}
    </Card>
  );
}

function TodayRow({ left, right, onClick, sub }) {
  return (
    <div onClick={onClick} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderTop: `1px solid ${S.line}`, fontSize: 13.5, cursor: onClick ? "pointer" : "default" }}>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontWeight: 600, color: onClick ? S.blue : S.ink, overflow: "hidden", textOverflow: "ellipsis" }}>{left}{onClick && " ›"}</div>
        {sub && <div style={{ fontSize: 12, color: S.sub }}>{sub}</div>}
      </div>
      {right && <div className="num" style={{ fontSize: 12.5, color: S.sub, textAlign: "right", flex: "0 0 auto" }}>{right}</div>}
    </div>
  );
}

function TodayTab({ data, api, admin, onOpenWorker, onOpenObject, openAdminPanel, goTab }) {
  const [closing, setClosing] = useState(false);
  const [remText, setRemText] = useState("");
  const [remDate, setRemDate] = useState("");
  const addRem = async () => { if (!remText.trim()) return; const ok = await api.addReminder(remText.trim(), remDate); if (ok) { setRemText(""); setRemDate(""); } };
  const today = todayISO();
  const prevMo = prevMonthKey(curMonth());
  const prevMoLabel = `${mName(Number(prevMo.slice(5, 7)) - 1)} ${prevMo.slice(0, 4)}.`;
  const wName = (id) => data.workers.find((w) => w.id === id)?.name || "Obrisan radnik";
  const oName = (id) => data.objects.find((o) => o.id === id)?.name || "bez objekta";
  const active = data.workers.filter((w) => !w.archived);
  const openObjs = data.objects.filter((o) => !o.archived);

  // podsjetnici
  const reminders = (data.reminders || []).filter((r) => !r.done);
  const dueRem = reminders.filter((r) => r.dueDate && r.dueDate <= today).sort((a, b) => a.dueDate.localeCompare(b.dueDate));
  const weekRem = reminders.filter((r) => r.dueDate && r.dueDate > today && r.dueDate <= addDaysISO(today, 7)).sort((a, b) => a.dueDate.localeCompare(b.dueDate));

  // dokumenti
  const warns = expiryWarnings(data.workers, data.settings?.expiry_warn_days);
  const noDocs = active.filter((w) => !w.permitExpiry && !w.contractExpiry);

  // isplate
  const pending = (data.payouts || []).filter((p) => !p.approved);
  const withHoursPrev = new Set(data.logs.filter((l) => monthKey(l.date) === prevMo).map((l) => l.workerId));
  const notProposedPrev = [...withHoursPrev].filter((id) => !paidFor(data, id, prevMo)).length;
  const paydayDay = Number(data.settings?.payday_day) || 5;
  const todayDay = new Date().getDate();
  const daysToPayday = paydayDay >= todayDay ? paydayDay - todayDay : null;

  // objekti bez sati za prošli mjesec (otvoreni objekti koji imaju aktivne radnike)
  const objHasHours = (o, mo) => data.logs.some((l) => l.objectId === o.id && monthKey(l.date) === mo)
    || (data.brigadeCosts || []).some((b) => b.objectId === o.id && monthKey(b.date) === mo);
  const missingHours = openObjs.filter((o) => active.some((w) => w.objectId === o.id) && !objHasHours(o, prevMo));

  // sumnjivi unosi (zadnja 4 mjeseca)
  const since = addDaysISO(today, -122);
  const recent = data.logs.filter((l) => l.date >= since);
  const suspicious = recent.filter((l) => (l.monthly ? l.hours > MONTH_HOURS_WARN : isDayHoursSuspicious(l.hours)))
    .sort((a, b) => b.date.localeCompare(a.date));
  const byDay = {};
  recent.filter((l) => !l.monthly).forEach((l) => { const k = l.workerId + "|" + l.date; (byDay[k] = byDay[k] || []).push(l); });
  const dups = Object.values(byDay).filter((a) => a.length > 1).sort((a, b) => b[0].date.localeCompare(a[0].date));

  // podaci koji nedostaju
  const noPhone = active.filter((w) => !(w.phone || "").trim());
  const noPos = active.filter((w) => !(w.position || "").trim());
  const noPrice = admin ? openObjs.filter((o) => !(o.billRate > 0) && !(data.positionBilling || []).some((pb) => pb.objectId === o.id)) : [];

  // DPH za prošli mjesec (admin)
  const vat = admin ? vatSummary(data, prevMo) : null;
  const vatDays = vat ? daysUntil(vat.due) : null;
  const showVat = admin && vat && (vat.out.length || vat.inn.length) && vatDays >= -10;

  const urgentCount = (showVat && vatDays <= 5 && vat.payable > 0 ? 1 : 0) + dueRem.length + warns.filter((w) => w.past).length + (admin ? pending.length : 0);
  const nothing = !admin && !showVat && !dueRem.length && !warns.length && !(admin && (pending.length || notProposedPrev)) && !missingHours.length && !suspicious.length && !dups.length;
  if (closing) return <MonthClose data={data} api={api} initialMonth={prevMo} onClose={() => setClosing(false)}
    onOpenObject={onOpenObject} onOpenWorker={onOpenWorker} goTab={goTab} openAdminPanel={openAdminPanel} />;
  const dateLabel = new Date().toLocaleDateString("hr-HR", { weekday: "long", day: "numeric", month: "numeric", year: "numeric" });

  return (
    <div>
      <Card style={{ background: urgentCount ? S.amberSoft : S.greenSoft, borderColor: urgentCount ? "#EBD9B4" : "#CFE3D8", padding: "12px 14px" }}>
        <div style={{ fontSize: 12.5, color: S.sub, textTransform: "capitalize" }}>{dateLabel}</div>
        <div style={{ fontWeight: 800, fontSize: 17, color: urgentCount ? S.amber : S.green, marginTop: 2 }}>
          {urgentCount ? tr("{0} {1} pažnju danas", urgentCount, urgentCount === 1 ? "stvar traži" : "stvari traže") : tr("Danas nema ništa hitno ✓")}
        </div>
        <div className="num" style={{ fontSize: 12.5, color: S.sub, marginTop: 4 }}>
          {active.length} {tr("aktivnih radnika ·")}{" "}{openObjs.length} {tr("otvorenih objekata")}
        </div>
      </Card>

      {admin && (
        <TodaySection icon="📌" title={tr("Podsjetnici")} count={dueRem.length ? `${dueRem.length} danas` : undefined} tone={dueRem.length ? "red" : "blue"}
          items={[...reminders].sort((x, y) => (x.dueDate || "9999").localeCompare(y.dueDate || "9999"))} limit={6}
          render={(r) => {
            const late = r.dueDate && r.dueDate < today, due = r.dueDate && r.dueDate <= today;
            return (
              <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderTop: `1px solid ${S.line}`, fontSize: 13.5 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: due ? 700 : 500 }}>{r.text}</div>
                  {r.dueDate && <div style={{ fontSize: 12, color: late ? S.red : due ? S.amber : S.sub }}>{late ? tr("kasni · ") : due ? tr("danas · ") : ""}{fmtDate(r.dueDate)}</div>}
                </div>
                <button onClick={() => api.toggleReminder(r, true)} title={tr("Gotovo")} style={{ background: S.greenSoft, border: "none", color: S.green, fontWeight: 800, borderRadius: 8, padding: "5px 10px", cursor: "pointer" }}>✓</button>
                <button onClick={() => api.delReminder(r)} title={tr("Obriši")} style={{ background: "none", border: "none", color: S.sub, fontSize: 14, cursor: "pointer", padding: 4 }}>✕</button>
              </div>
            );
          }}>
          {reminders.length === 0 && <div style={{ fontSize: 13, color: S.sub, marginBottom: 6 }}>{tr("Nema otvorenih podsjetnika.")}</div>}
          <div style={{ display: "flex", gap: 6, marginBottom: reminders.length ? 6 : 0 }}>
            <input value={remText} onChange={(e) => setRemText(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addRem()} placeholder={tr("Novi podsjetnik…")} style={{ flex: 2, padding: "9px 10px" }} />
            <input type="date" value={remDate} onChange={(e) => setRemDate(e.target.value)} style={{ flex: 1, padding: "9px 8px" }} />
            <Btn small onClick={addRem} disabled={!remText.trim()}>+</Btn>
          </div>
        </TodaySection>
      )}

      {api.recruiter && (data.newCandidates > 0 || !admin) && (
        <TodaySection icon="🧑‍🍳" title={tr("Kandidati")} tone={data.staleCandidates ? "red" : data.newCandidates ? "amber" : "blue"} count={data.newCandidates ? `${data.newCandidates} ${tr("novih")}` : undefined}
          action={<Btn small onClick={() => { goTab("kandidati"); window.scrollTo({ top: 0 }); }} style={{ width: "100%" }}>{tr("Otvori kandidate")}</Btn>}>
          <div style={{ fontSize: 13.5 }}>{data.newCandidates ? tr("Nove prijave s prijavnice čekaju da im se javiš.") : tr("Nema novih prijava.")}</div>
          {data.staleCandidates > 0 && <div style={{ fontSize: 13.5, color: S.red, fontWeight: 600, marginTop: 4 }}>⏰ {data.staleCandidates} {tr("čeka dulje od 24 h — javi im se danas.")}</div>}
        </TodaySection>
      )}

      {admin && (pending.length > 0 || notProposedPrev > 0 || daysToPayday !== null) && (
        <TodaySection icon="💶" title={tr("Isplate")} tone={pending.length ? "amber" : "blue"} count={pending.length ? `${pending.length} čeka` : undefined}
          action={pending.length > 0 && <Btn small onClick={() => openAdminPanel("approvals")} style={{ width: "100%" }}>{tr("Otvori odobravanje isplata")}</Btn>}>
          <div style={{ fontSize: 13.5, display: "grid", gap: 4 }}>
            {daysToPayday !== null && (
              <div>📅 {daysToPayday === 0 ? <b>{tr("Danas je dan isplate.")}</b> : <>{tr("Dan isplate je za")}{" "}<b>{daysToPayday} {daysToPayday === 1 ? tr("dan") : tr("dana")}</b> ({paydayDay}{tr(". u mjesecu).")}</>}</div>
            )}
            {pending.length > 0 && <div>⏳ <b>{pending.length}</b> {tr("predloženih isplata čeka odobrenje.")}</div>}
            {notProposedPrev > 0 && (
              <div>{tr("📝 Za")}{" "}{prevMoLabel} {tr("još")}{" "}<b>{notProposedPrev}</b> {tr("radnika sa satima nema predloženu isplatu.")}{" "}
                <span onClick={() => goTab("obracun")} style={{ color: S.blue, fontWeight: 700, cursor: "pointer" }}>{tr("Obračun ›")}</span></div>
            )}
          </div>
        </TodaySection>
      )}

      {admin && <MonthCloseCard data={data} mo={prevMo} onStart={() => setClosing(true)} />}

      {showVat && (
        <TodaySection icon="🧾" title={tr("DPH za {0}", prevMoLabel)} tone={vatDays < 0 ? "red" : vatDays <= 5 ? "amber" : "blue"}
          count={vat.payable >= 0 ? czk(vat.payable) : "vraća se " + czk(-vat.payable)}
          action={<Btn small kind="ghost" onClick={() => openAdminPanel("vat")} style={{ width: "100%" }}>{tr("Detalji DPH")}</Btn>}>
          <div style={{ fontSize: 13.5 }}>
            {tr("Rok za prijavu i plaćanje:")}{" "}<b>{fmtDate(vat.due)}</b>{" "}
            <span style={{ color: vatDays < 0 ? S.red : S.sub }}>({vatDays === 0 ? tr("danas") : vatDays > 0 ? tr("za {0} {1}", vatDays, vatDays === 1 ? "dan" : "dana") : tr("kasni {0} {1}", -vatDays, -vatDays === 1 ? "dan" : "dana")})</span>
            <div style={{ fontSize: 12, color: S.sub, marginTop: 2 }}>{vat.out.length} {tr("izdanih faktura · DPH")}{" "}{czk(vat.outVat)}{vat.inVat ? tr(" − odbitak {0}", czk(vat.inVat)) : ""}</div>
          </div>
        </TodaySection>
      )}

      {missingHours.length > 0 && (
        <TodaySection icon="🕒" title={tr("Objekti bez upisanih sati za {0}", prevMoLabel)} count={missingHours.length} tone="amber" items={missingHours}
          render={(o) => (
            <TodayRow key={o.id} left={`${FLAG[o.country] || ""} ${o.name}`} onClick={() => onOpenObject(o.id)}
              right={`${active.filter((w) => w.objectId === o.id).length} radnika`} />
          )} />
      )}

      {warns.length > 0 && (
        <TodaySection icon="🪪" title={tr("Dozvole i ugovori koji ističu")} count={warns.length} tone={warns.some((w) => w.past) ? "red" : "amber"} items={warns}
          render={(x, i) => (
            <TodayRow key={x.w.id + x.label + i} left={x.w.name} sub={`${x.label} · ${oName(x.w.objectId)}`} onClick={() => onOpenWorker(x.w.id)}
              right={<span style={{ color: x.past ? S.red : S.amber, fontWeight: 700 }}>{x.past ? tr("isteklo ") : ""}{fmtDate(x.date)}</span>} />
          )} />
      )}

      {(suspicious.length > 0 || dups.length > 0) && (
        <TodaySection icon="🔎" title={tr("Unosi sati za provjeru")} count={suspicious.length + dups.length} tone="amber"
          items={[...suspicious.map((l) => ({ kind: "big", l })), ...dups.map((a) => ({ kind: "dup", a }))]}
          render={(x) => x.kind === "big" ? (
            <TodayRow key={"b" + x.l.id} left={wName(x.l.workerId)} onClick={() => onOpenWorker(x.l.workerId)}
              sub={`${x.l.monthly ? "mjesečni zbroj iznad " + MONTH_HOURS_WARN + " h" : "jedan dan, više od " + DAY_HOURS_WARN + " h — možda je trebao biti mjesečni zbroj"} · ${oName(x.l.objectId)}`}
              right={<><b style={{ color: S.amber }}>{fmtH(x.l.hours)}</b><br />{x.l.monthly ? x.l.date.slice(0, 7) : fmtDate(x.l.date)}</>} />
          ) : (
            <TodayRow key={"d" + x.a[0].id} left={wName(x.a[0].workerId)} onClick={() => onOpenWorker(x.a[0].workerId)}
              sub={tr("{0} unosa za isti dan ({1}) — možda duplikat", x.a.length, x.a.map((l) => fmtH(l.hours)).join(" + "))}
              right={fmtDate(x.a[0].date)} />
          )}>
          <div style={{ fontSize: 12, color: S.sub, marginBottom: 4 }}>{tr("Otvori radnika i ispravi ili obriši unos ako je greška.")}</div>
        </TodaySection>
      )}

      {(noDocs.length > 0 || noPhone.length > 0 || noPos.length > 0 || noPrice.length > 0) && (
        <TodaySection icon="🧩" title={tr("Podaci koji nedostaju")} tone="blue">
          <div style={{ fontSize: 12, color: S.sub, marginBottom: 6 }}>{tr("Bez ovoga ne rade upozorenja o isteku, kontakt ni naplata po poziciji.")}</div>
          {[
            [noDocs, "bez datuma isteka dozvole/ugovora"],
            [noPhone, "bez broja telefona"],
            [noPos, "bez pozicije"],
          ].filter(([list]) => list.length).map(([list, label]) => (
            <MissingList key={label} list={list} label={label} oName={oName} onOpenWorker={onOpenWorker} />
          ))}
          {noPrice.length > 0 && (
            <MissingList list={noPrice} label={tr("objekata bez cijene naplate")} oName={() => ""} onOpenWorker={onOpenObject} isObject />
          )}
        </TodaySection>
      )}

      {nothing && <Empty text={tr("Sve je pod kontrolom — nema zakašnjelih podsjetnika, isteka ni sumnjivih unosa. 🎉")} />}
    </div>
  );
}

function MissingList({ list, label, oName, onOpenWorker, isObject }) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ borderTop: `1px solid ${S.line}`, padding: "7px 0" }}>
      <div onClick={() => setOpen(!open)} style={{ display: "flex", justifyContent: "space-between", cursor: "pointer", fontSize: 13.5 }}>
        <span><b className="num">{list.length}</b> {isObject ? label : tr("radnika {0}", label)}</span>
        <span style={{ color: S.blue, fontWeight: 700 }}>{open ? "▲" : "▼"}</span>
      </div>
      {open && (
        <div style={{ maxHeight: 260, overflowY: "auto", marginTop: 4 }}>
          {list.map((x) => (
            <div key={x.id} onClick={() => onOpenWorker(x.id)} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0 5px 8px", fontSize: 13, cursor: "pointer" }}>
              <span style={{ color: S.blue, fontWeight: 600 }}>{isObject ? `${FLAG[x.country] || ""} ${x.name}` : x.name} ›</span>
              {!isObject && <span style={{ color: S.sub, fontSize: 12 }}>{oName(x.objectId)}</span>}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ================================================================== */
/*  📥 UVOZ SATI IZ EXCELA HOTELA — svaki hotel ima svoj format         */
/* ================================================================== */
const foldText = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "")
  .replace(/[^a-z\s]/g, " ").replace(/\s+/g, " ").trim();
const nameKeyOf = (s) => foldText(s).split(" ").filter(Boolean).sort().join(" ");
const TOTAL_ROW_RE = /^(celkem|ukupno|total|suma|soucet|zbroj|sum)\b/;
function parseHoursCell(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const t = String(v).trim();
  const hm = t.match(/^(\d{1,3}):(\d{2})(?::\d{2})?$/);
  if (hm) return Number(hm[1]) + Number(hm[2]) / 60;
  const n = parseFloat(t.replace(/\s/g, "").replace(",", "."));
  return isNaN(n) ? 0 : n;
}
const colLetter = (i) => { let s = ""; i += 1; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
const lsGet = (k, fb) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : fb; } catch { return fb; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} };

// pronađi radnika za ime iz Excela: zapamćeni par → isto ime (bilo koji redoslijed) → najviše zajedničkih riječi
function matchWorkerName(excelName, workers, preferIds, aliases) {
  const f = foldText(excelName);
  if (aliases[f]) { const w = workers.find((x) => x.id === aliases[f]); if (w) return { w, how: "zapamćeno" }; }
  const key = nameKeyOf(excelName);
  const exact = workers.filter((w) => nameKeyOf(w.name) === key);
  if (exact.length === 1) return { w: exact[0], how: "točno" };
  if (exact.length > 1) { const p = exact.find((w) => preferIds.has(w.id)); if (p) return { w: p, how: "točno" }; }
  const toks = new Set(f.split(" ").filter((t) => t.length >= 2));
  if (!toks.size) return null;
  let best = [];
  let bestScore = 0;
  workers.forEach((w) => {
    const wt = foldText(w.name).split(" ").filter((t) => t.length >= 2);
    const shared = wt.filter((t) => toks.has(t)).length;
    if (!shared) return;
    const score = shared / Math.max(wt.length, toks.size);
    if (score > bestScore + 0.001) { bestScore = score; best = [w]; } else if (Math.abs(score - bestScore) <= 0.001) best.push(w);
  });
  if (bestScore >= 0.5 && best.length === 1) return { w: best[0], how: "slično" };
  return null;
}

function guessLayout(rows, workers) {
  const scan = rows.slice(0, 20);
  let headerRow = scan.findIndex((r) => r.filter((c) => String(c).trim() && isNaN(parseHoursCell(c) || NaN)).length >= 2);
  if (headerRow < 0) headerRow = 0;
  const body = rows.slice(headerRow + 1, headerRow + 200);
  const width = Math.max(0, ...rows.slice(0, 200).map((r) => r.length));
  const wKeys = workers.map((w) => foldText(w.name).split(" "));
  // stupac s imenima: najviše ćelija koje dijele riječ s nekim radnikom
  let nameCol = 0, nameScore = -1;
  for (let c = 0; c < width; c++) {
    let sc = 0;
    body.forEach((r) => {
      const t = foldText(r[c]).split(" ").filter((x) => x.length >= 3);
      if (t.length && wKeys.some((wk) => t.some((x) => wk.includes(x)))) sc += 2;
      else if (t.length >= 2) sc += 0.5;
    });
    if (sc > nameScore) { nameScore = sc; nameCol = c; }
  }
  const header = (rows[headerRow] || []).map((h) => foldText(h));
  const numericShare = (c) => body.filter((r) => parseHoursCell(r[c]) > 0).length / Math.max(1, body.length);
  // dani 1..31 u naslovima → zbrajanje stupaca
  const dayCols = header.map((h, i) => (/^\d{1,2}$/.test(h.trim()) || /^\d{1,2}$/.test(String(rows[headerRow]?.[i] ?? "").trim().replace(/\.$/, "")) ? i : -1))
    .filter((i) => i >= 0 && i !== nameCol);
  let hoursCol = header.findIndex((h, i) => i !== nameCol && /(celkem|ukupno|total|suma|hodin|sati|hours|odpracovano)/.test(h) && numericShare(i) > 0.2);
  const namedTotal = hoursCol >= 0;
  if (hoursCol < 0) {
    for (let c = width - 1; c >= 0; c--) if (c !== nameCol && numericShare(c) > 0.3) { hoursCol = c; break; }
  }
  const mode = !namedTotal && dayCols.length >= 5 ? "sum" : "col";
  return { headerRow, nameCol, hoursCol: Math.max(0, hoursCol), mode, fromCol: dayCols[0] ?? 0, toCol: dayCols[dayCols.length - 1] ?? 0 };
}

function HoursImport({ data, object, month, onSave, onClose }) {
  const mapKey = "evidencija_import_map_" + object.id;
  const aliasKey = "evidencija_import_alias_" + object.id;
  const [wb, setWb] = useState(null);
  const [fileName, setFileName] = useState("");
  const [sheetName, setSheetName] = useState("");
  const [layout, setLayout] = useState(null);
  const [overrides, setOverrides] = useState({}); // redak → workerId ("" = preskoči)
  const [err, setErr] = useState("");
  const [saving, setSaving] = useState(false);

  const active = sortedWorkers(data.workers.filter((w) => !w.archived));
  const atObject = workersAtObject(data, object.id);
  const preferIds = new Set(atObject.map((w) => w.id));
  const options = [...atObject, ...active.filter((w) => !preferIds.has(w.id))];
  const aliases = lsGet(aliasKey, {});

  const rows = wb && sheetName ? XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: false, defval: "" }) : [];
  const width = Math.max(0, ...rows.slice(0, 200).map((r) => r.length));
  const headerCells = layout ? (rows[layout.headerRow] || []) : [];
  const colName = (i) => `${colLetter(i)}${headerCells[i] ? " – " + String(headerCells[i]).slice(0, 24) : ""}`;

  const loadFile = async (file) => {
    setErr("");
    try {
      const buf = await file.arrayBuffer();
      const book = XLSX.read(buf, { type: "array" });
      setWb(book); setFileName(file.name); setOverrides({});
      const saved = lsGet(mapKey, null);
      const sn = saved && book.SheetNames.includes(saved.sheetName) ? saved.sheetName : book.SheetNames[0];
      setSheetName(sn);
      const r = XLSX.utils.sheet_to_json(book.Sheets[sn], { header: 1, raw: false, defval: "" });
      setLayout(saved && saved.sheetName === sn ? { ...guessLayout(r, active), ...saved.layout } : guessLayout(r, active));
    } catch (e) { setErr(tr("Ne mogu pročitati datoteku: ") + (e.message || e)); }
  };
  const changeSheet = (sn) => {
    setSheetName(sn); setOverrides({});
    setLayout(guessLayout(XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: false, defval: "" }), active));
  };
  const setL = (patch) => { setLayout((l) => ({ ...l, ...patch })); setOverrides({}); };

  // redovi iz Excela → ime, sati, pronađeni radnik
  const parsed = [];
  if (layout) {
    rows.slice(layout.headerRow + 1).forEach((r, i) => {
      const rowIdx = layout.headerRow + 1 + i;
      const name = String(r[layout.nameCol] ?? "").trim();
      if (!name || !/[a-zA-Zà-žА-я]/.test(name) || TOTAL_ROW_RE.test(foldText(name))) return;
      let hours = 0;
      if (layout.mode === "sum") {
        const a = Math.min(layout.fromCol, layout.toCol), b = Math.max(layout.fromCol, layout.toCol);
        for (let c = a; c <= b; c++) hours += parseHoursCell(r[c]);
      } else hours = parseHoursCell(r[layout.hoursCol]);
      hours = round2(hours);
      if (!(hours > 0)) return;
      const auto = matchWorkerName(name, active, preferIds, aliases);
      const chosenId = overrides[rowIdx] !== undefined ? overrides[rowIdx] : (auto?.w.id || "");
      parsed.push({ rowIdx, name, hours, auto, workerId: chosenId });
    });
  }
  const byWorker = {};
  parsed.filter((p) => p.workerId).forEach((p) => { byWorker[p.workerId] = round2((byWorker[p.workerId] || 0) + p.hours); });
  const countOf = (wid) => parsed.filter((p) => p.workerId === wid).length;
  const current = (wid) => round2(data.logs.filter((l) => l.workerId === wid && l.objectId === object.id && l.monthly && monthKey(l.date) === month).reduce((s, l) => s + l.hours, 0));
  const matched = Object.keys(byWorker);
  const unmatched = parsed.filter((p) => !p.workerId);
  const totalHours = round2(matched.reduce((s, id) => s + byWorker[id], 0));

  const save = async () => {
    const entries = matched.map((wid) => ({ workerId: wid, hours: byWorker[wid] }));
    const tooMany = entries.filter((e) => e.hours > MONTH_HOURS_WARN);
    if (tooMany.length && !window.confirm(tr("Više od {0} h ima {1} radnik(a). Provjeri jesi li odabrala pravi stupac sa satima.\n\nSpremiti svejedno?", MONTH_HOURS_WARN, tooMany.length))) return;
    if (!window.confirm(tr("Upisati mjesečne sate za {0} radnika ({1} ukupno) za {2} {3}. na {4}?\n\nPostojeći mjesečni zbroj tih radnika za taj mjesec bit će zamijenjen brojem iz Excela.", entries.length, fmtH(totalHours), MONTHS[Number(month.slice(5, 7)) - 1], month.slice(0, 4), object.name))) return;
    // zapamti format i ručno odabrane parove za ovaj hotel
    lsSet(mapKey, { sheetName, layout });
    const al = { ...aliases };
    parsed.forEach((p) => { if (overrides[p.rowIdx] !== undefined) { if (p.workerId) al[foldText(p.name)] = p.workerId; else delete al[foldText(p.name)]; } });
    lsSet(aliasKey, al);
    setSaving(true);
    const ok = await onSave(entries);
    setSaving(false);
    if (ok !== false) onClose();
  };

  const sel = { padding: "7px 8px", fontSize: 13 };
  return (
    <Card style={{ border: `2px solid ${S.green}` }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <div style={{ fontWeight: 800 }}>{tr("📥 Uvoz sati iz Excela ·")}{" "}{mName(Number(month.slice(5, 7)) - 1)} {month.slice(0, 4)}.</div>
        <button onClick={onClose} style={{ background: "none", border: "none", fontWeight: 800, color: S.sub, cursor: "pointer", fontSize: 16 }}>✕</button>
      </div>
      <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 10 }}>
        {tr("Učitaj tablicu koju je poslao hotel. Aplikacija pogodi gdje su imena i sati, a ti samo provjeriš. Format i ručno spojena imena pamte se za ovaj hotel (na ovom uređaju), pa idući put ide brže.")}
      </div>
      <input type="file" accept=".xlsx,.xls,.xlsm,.csv,.ods" onChange={(e) => e.target.files[0] && loadFile(e.target.files[0])} />
      {err && <div style={{ color: S.red, fontSize: 13, marginTop: 6 }}>{err}</div>}

      {wb && layout && (
        <>
          <div style={{ fontSize: 12, color: S.sub, margin: "6px 0 10px" }}>{fileName}</div>
          {wb.SheetNames.length > 1 && (
            <Field label={tr("List u tablici")}>
              <select value={sheetName} onChange={(e) => changeSheet(e.target.value)} style={sel}>
                {wb.SheetNames.map((n) => <option key={n} value={n}>{n}</option>)}
              </select>
            </Field>
          )}
          <div style={{ display: "flex", gap: 8 }}>
            <div style={{ flex: 1 }}>
              <Field label={tr("Redak s naslovima")}>
                <select value={layout.headerRow} onChange={(e) => setL({ headerRow: Number(e.target.value) })} style={sel}>
                  {rows.slice(0, 25).map((r, i) => <option key={i} value={i}>{i + 1}. {r.filter((c) => String(c).trim()).slice(0, 3).join(" | ").slice(0, 30)}</option>)}
                </select>
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label={tr("Stupac s imenom")}>
                <select value={layout.nameCol} onChange={(e) => setL({ nameCol: Number(e.target.value) })} style={sel}>
                  {Array.from({ length: width }, (_, i) => <option key={i} value={i}>{colName(i)}</option>)}
                </select>
              </Field>
            </div>
          </div>
          <Field label={tr("Sati")}>
            <div style={{ display: "flex", gap: 14, fontSize: 13, marginBottom: 6 }}>
              <label style={{ display: "flex", gap: 5, alignItems: "center", cursor: "pointer" }}>
                <input type="radio" checked={layout.mode === "col"} onChange={() => setL({ mode: "col" })} style={{ width: 16, height: 16 }} /> {tr("jedan stupac s ukupnim satima")}
              </label>
              <label style={{ display: "flex", gap: 5, alignItems: "center", cursor: "pointer" }}>
                <input type="radio" checked={layout.mode === "sum"} onChange={() => setL({ mode: "sum" })} style={{ width: 16, height: 16 }} /> {tr("zbroji stupce (po danima)")}
              </label>
            </div>
            {layout.mode === "col" ? (
              <select value={layout.hoursCol} onChange={(e) => setL({ hoursCol: Number(e.target.value) })} style={sel}>
                {Array.from({ length: width }, (_, i) => <option key={i} value={i}>{colName(i)}</option>)}
              </select>
            ) : (
              <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <select value={layout.fromCol} onChange={(e) => setL({ fromCol: Number(e.target.value) })} style={sel}>
                  {Array.from({ length: width }, (_, i) => <option key={i} value={i}>{tr("od")}{" "}{colName(i)}</option>)}
                </select>
                <select value={layout.toCol} onChange={(e) => setL({ toCol: Number(e.target.value) })} style={sel}>
                  {Array.from({ length: width }, (_, i) => <option key={i} value={i}>{tr("do")}{" "}{colName(i)}</option>)}
                </select>
              </div>
            )}
          </Field>

          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", margin: "4px 0 8px" }}>
            <Tag color={S.green} bg={S.greenSoft}>✓ {matched.length} {tr("radnika ·")}{" "}{fmtH(totalHours)}</Tag>
            {unmatched.length > 0 && <Tag color={S.amber} bg={S.amberSoft}>? {unmatched.length} {tr("nije prepoznato")}</Tag>}
          </div>
          {parsed.length === 0 && <div style={{ fontSize: 13, color: S.sub, marginBottom: 8 }}>{tr("Nema redaka s imenom i satima — promijeni redak s naslovima ili stupce iznad.")}</div>}

          <div style={{ maxHeight: 380, overflowY: "auto", borderTop: `1px solid ${S.line}` }}>
            {[...unmatched, ...parsed.filter((p) => p.workerId)].map((p) => {
              const cur = p.workerId ? current(p.workerId) : 0;
              return (
                <div key={p.rowIdx} style={{ padding: "7px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
                    <span style={{ fontWeight: 600 }}>{p.name} <span style={{ color: S.sub, fontWeight: 400, fontSize: 11.5 }}>{tr("(red")}{" "}{p.rowIdx + 1})</span></span>
                    <b className="num">{fmtH(p.hours)}</b>
                  </div>
                  <div style={{ display: "flex", gap: 6, alignItems: "center", marginTop: 4 }}>
                    <select value={p.workerId} onChange={(e) => setOverrides((o) => ({ ...o, [p.rowIdx]: e.target.value }))}
                      style={{ ...sel, flex: 1, borderColor: p.workerId ? S.line : S.amber, color: p.workerId ? S.ink : S.amber }}>
                      <option value="">{tr("— nije prepoznat, preskoči —")}</option>
                      {options.map((w) => <option key={w.id} value={w.id}>{w.name}{preferIds.has(w.id) ? "" : tr(" (drugi objekt)")}</option>)}
                    </select>
                  </div>
                  {p.workerId && (
                    <div style={{ fontSize: 11.5, color: S.sub, marginTop: 3 }}>
                      {p.auto && p.auto.w.id === p.workerId ? tr("prepoznato: {0}", p.auto.how) : tr("ručno odabrano")}
                      {countOf(p.workerId) > 1 && <b style={{ color: S.amber }}> · {countOf(p.workerId)} {tr("retka zbrojena")}</b>}
                      {cur > 0 && <> {tr("· sad u aplikaciji")}{" "}{fmtH(cur)} {tr("→ bit će")}{" "}{fmtH(byWorker[p.workerId])}</>}
                    </div>
                  )}
                </div>
              );
            })}
          </div>

          <Btn onClick={save} disabled={saving || matched.length === 0} style={{ width: "100%", marginTop: 10 }}>
            {saving ? tr("Spremam…") : tr("💾 Upiši sate za {0} radnika", matched.length)}
          </Btn>
        </>
      )}
    </Card>
  );
}

function AddEmployee({ data, api }) {
  const [open, setOpen] = useState(false);
  const [f, setF] = useState({ name: "", email: "", password: "" });
  const [objs, setObjs] = useState(() => new Set());
  const [done, setDone] = useState(null);
  const genPass = () => {
    const chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789";
    const a = new Uint32Array(10); crypto.getRandomValues(a);
    setF((p) => ({ ...p, password: Array.from(a, (x) => chars[x % chars.length]).join("") }));
  };
  const submit = async () => {
    const ok = await api.createEmployee({ name: normalizeName(f.name), email: f.email.trim(), password: f.password, objectIds: [...objs] });
    if (ok) { setDone({ ...f }); setF({ name: "", email: "", password: "" }); setObjs(new Set()); }
  };
  const openObjs = sortedObjects(data.objects.filter((o) => !o.archived));
  return (
    <div style={{ marginTop: 10, borderTop: `1px solid ${S.line}`, paddingTop: 10 }}>
      {!open ? (
        <Btn small onClick={() => { setOpen(true); setDone(null); }} style={{ width: "100%" }}>{tr("➕ Dodaj zaposlenika (koordinatora)")}</Btn>
      ) : (
        <div>
          <div style={{ fontWeight: 700, marginBottom: 6 }}>{tr("➕ Novi zaposlenik")}</div>
          <div style={{ fontSize: 12, color: S.sub, marginBottom: 8 }}>
            {tr("Zaposlenik vidi i upisuje samo objekte koje mu dodijeliš. Ne vidi cijene naplate hotelima, fakture, maržu ni tuđe provizije.")}
          </div>
          <Field label={tr("Ime i prezime")}><input value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} placeholder={tr("npr. Maryna Svydenyuk")} /></Field>
          <Field label={tr("E-mail za prijavu")}><input type="email" value={f.email} onChange={(e) => setF({ ...f, email: e.target.value })} placeholder={tr("ime@email.com")} /></Field>
          <Field label={tr("Privremena lozinka (barem 8 znakova)")}>
            <div style={{ display: "flex", gap: 6 }}>
              <input value={f.password} onChange={(e) => setF({ ...f, password: e.target.value })} />
              <Btn small kind="ghost" onClick={genPass}>{tr("🎲 Izmisli")}</Btn>
            </div>
          </Field>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: S.sub, marginBottom: 4 }}>{tr("Objekti koje vodi")}</div>
          <div style={{ maxHeight: 200, overflowY: "auto", marginBottom: 10 }}>
            {openObjs.map((o) => (
              <label key={o.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "4px 0", fontSize: 13.5, cursor: "pointer" }}>
                <input type="checkbox" checked={objs.has(o.id)} style={{ width: 17, height: 17 }}
                  onChange={(e) => setObjs((p) => { const n = new Set(p); e.target.checked ? n.add(o.id) : n.delete(o.id); return n; })} />
                {FLAG[o.country] || ""} {o.name}
              </label>
            ))}
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <Btn small kind="ghost" onClick={() => setOpen(false)} style={{ flex: 1 }}>{tr("Odustani")}</Btn>
            <Btn small onClick={submit} disabled={!f.name.trim() || !f.email.includes("@") || f.password.length < 8} style={{ flex: 2 }}>{tr("✓ Kreiraj račun")}</Btn>
          </div>
        </div>
      )}
      {done && (
        <div style={{ marginTop: 10, background: S.greenSoft, borderRadius: 10, padding: "9px 12px", fontSize: 13 }}>
          {tr("✓ Račun za")}{" "}<b>{done.name}</b> {tr("je kreiran. Pošalji joj/mu podatke za prijavu na evidecija.vercel.app:")}<br />
          {tr("e-mail")}{" "}<b>{done.email}</b> {tr("· lozinka")}{" "}<b style={{ fontFamily: "monospace" }}>{done.password}</b><br />
          <span style={{ color: S.sub }}>{tr("Lozinka se ovdje prikazuje samo sada.")}</span>
        </div>
      )}
    </div>
  );
}

/* ================================================================== */
/*  ✅ ZATVARANJE MJESECA — vodi kroz svaki objekt korak po korak      */
/* ================================================================== */
function monthCloseObjects(data, mo) {
  const active = data.workers.filter((w) => !w.archived);
  return sortedObjects(data.objects.filter((o) => {
    const hasLogs = data.logs.some((l) => l.objectId === o.id && monthKey(l.date) === mo);
    const hasInv = (data.objectInvoices || []).some((i) => i.objectId === o.id && (invoiceTaxMonth(i) === mo || i.period === mo));
    const hasWorkers = !o.archived && active.some((w) => w.objectId === o.id);
    return hasLogs || hasInv || hasWorkers;
  }));
}
function objectCloseState(data, o, mo) {
  const rate = Number(data.settings?.czk_rate) || 25;
  const logs = data.logs.filter((l) => l.objectId === o.id && monthKey(l.date) === mo);
  const brigade = (data.brigadeCosts || []).filter((b) => b.objectId === o.id && monthKey(b.date) === mo);
  const hours = round2(logs.reduce((s, l) => s + l.hours, 0) + brigade.reduce((s, b) => s + b.hours, 0));
  const withHours = new Set(logs.map((l) => l.workerId));
  const assignedNoHours = data.workers.filter((w) => !w.archived && w.objectId === o.id && !withHours.has(w.id));
  const suspicious = logs.filter((l) => (l.monthly ? isHoursSuspicious(l.hours) : isDayHoursSuspicious(l.hours)));
  const byDay = {};
  logs.filter((l) => !l.monthly).forEach((l) => { const k = l.workerId + "|" + l.date; byDay[k] = (byDay[k] || 0) + 1; });
  const dups = Object.values(byDay).filter((n) => n > 1).length;
  const invoices = (data.objectInvoices || []).filter((i) => i.objectId === o.id && (invoiceTaxMonth(i) === mo || i.period === mo));
  const invNetCzk = round2(invoices.reduce((s, i) => s + toCzk(i.amount - (i.vatAmount || 0), i.currency, rate), 0));
  const rv = objRevenue(data, o, logs, brigade);
  const expectedCzk = round2(rv.revenueKc + rv.revenue * rate);
  const diffPct = expectedCzk > 0 && invNetCzk > 0 ? round2(((invNetCzk - expectedCzk) / expectedCzk) * 100) : null;
  const paid = round2((data.invoicePayments || []).filter((p) => p.objectId === o.id && p.month === mo).reduce((s, p) => s + p.amount, 0));
  return { logs, hours, withHours, assignedNoHours, suspicious, dups, invoices, invNetCzk, expectedCzk, diffPct, paid };
}

function MonthCloseCard({ data, mo, onStart }) {
  const objs = monthCloseObjects(data, mo);
  const done = objs.filter((o) => (data.monthClose || []).some((c) => c.month === mo && c.objectId === o.id && c.step === "done")).length;
  const all = objs.length > 0 && done === objs.length;
  const label = `${MONTHS[Number(mo.slice(5, 7)) - 1]} ${mo.slice(0, 4)}.`;
  return (
    <TodaySection icon="✅" title={tr("Zatvaranje mjeseca: {0}", label)} tone={all ? "green" : "blue"} count={`${done}/${objs.length}`}
      action={<Btn small onClick={onStart} style={{ width: "100%" }}>{done === 0 ? tr("▶ Započni zatvaranje") : all ? tr("Pregledaj") : tr("▶ Nastavi")}</Btn>}>
      <div style={{ fontSize: 12.5, color: S.sub }}>
        {all ? tr("Svi objekti su zatvoreni za ovaj mjesec. ✓") : tr("Vodi te kroz svaki objekt: sati, provjere, faktura, pa na kraju isplate i DPH.")}
      </div>
    </TodaySection>
  );
}

function MonthClose({ data, api, initialMonth, onClose, onOpenObject, onOpenWorker, goTab, openAdminPanel }) {
  const [mo, setMo] = useState(initialMonth);
  const objs = monthCloseObjects(data, mo);
  const [idx, setIdx] = useState(0);
  const steps = [...objs.map((o) => ({ kind: "object", o })), { kind: "payouts" }, { kind: "vat" }, { kind: "summary" }];
  const cur = steps[Math.min(idx, steps.length - 1)];
  const label = `${MONTHS[Number(mo.slice(5, 7)) - 1]} ${mo.slice(0, 4)}.`;
  const isDone = (objectId, step) => (data.monthClose || []).some((c) => c.month === mo && (c.objectId || null) === (objectId || null) && c.step === step);
  const toggle = (objectId, step, name) => api.setCloseStep(mo, objectId, step, !isDone(objectId, step), name);
  const shiftMo = (n) => { const [y, m] = mo.split("-").map(Number); const d = new Date(y, m - 1 + n, 1); setMo(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`); setIdx(0); };
  const wName = (id) => data.workers.find((w) => w.id === id)?.name || "—";

  const Check = ({ ok, warn, children }) => (
    <div style={{ display: "flex", gap: 8, padding: "8px 0", borderTop: `1px solid ${S.line}`, fontSize: 13.5 }}>
      <span style={{ fontSize: 16, lineHeight: "20px" }}>{ok ? "✅" : warn ? "⚠️" : "⬜"}</span>
      <div style={{ flex: 1 }}>{children}</div>
    </div>
  );
  const Sub = ({ children }) => <div style={{ fontSize: 12, color: S.sub, marginTop: 2 }}>{children}</div>;
  const Link = ({ onClick, children }) => <span onClick={onClick} style={{ color: S.blue, fontWeight: 700, cursor: "pointer" }}>{children}</span>;

  let body = null;
  if (cur.kind === "object") {
    const o = cur.o;
    const st = objectCloseState(data, o, mo);
    const hoursOk = st.hours > 0 && (st.assignedNoHours.length === 0 || isDone(o.id, "hours"));
    const checksOk = st.suspicious.length === 0 && st.dups === 0;
    const invOk = st.invoices.length > 0;
    const done = isDone(o.id, "done");
    body = (
      <>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <div style={{ fontWeight: 800, fontSize: 16 }}>{FLAG[o.country] || ""} {o.name}</div>
          <Btn small kind="ghost" onClick={() => onOpenObject(o.id)}>{tr("Otvori objekt ›")}</Btn>
        </div>

        <Check ok={hoursOk} warn={st.hours > 0 && !hoursOk}>
          <b>{tr("1. Sati upisani")}</b> — {fmtH(st.hours)}, {st.withHours.size} {tr("radnika")}
          {st.hours === 0 && <Sub>{tr("Za ovaj mjesec nema sati. Upiši ih ili uvezi Excel hotela u objektu (Mjesečno → 📥 Uvezi sate).")}</Sub>}
          {st.assignedNoHours.length > 0 && (
            <Sub>
              {tr("Bez sati (")}{st.assignedNoHours.length}): {st.assignedNoHours.slice(0, 8).map((w, i) => (
                <span key={w.id}>{i ? ", " : ""}<Link onClick={() => onOpenWorker(w.id)}>{w.name}</Link></span>
              ))}{st.assignedNoHours.length > 8 ? " …" : ""}
              <label style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 4, cursor: "pointer", color: S.ink }}>
                <input type="checkbox" checked={isDone(o.id, "hours")} onChange={() => toggle(o.id, "hours", o.name)} style={{ width: 16, height: 16 }} />
                {tr("U redu je — ti radnici ovaj mjesec nisu radili ovdje")}
              </label>
            </Sub>
          )}
        </Check>

        <Check ok={checksOk} warn={!checksOk}>
          <b>{tr("2. Provjera unosa")}</b> — {checksOk ? tr("nema sumnjivih ni duplih unosa") : tr("{0} sumnjivih, {1} duplih dana", st.suspicious.length, st.dups)}
          {st.suspicious.slice(0, 5).map((l) => (
            <Sub key={l.id}><Link onClick={() => onOpenWorker(l.workerId)}>{wName(l.workerId)}</Link>: {fmtH(l.hours)} {l.monthly ? tr("(mjesečno)") : tr("u jednom danu")}</Sub>
          ))}
        </Check>

        <Check ok={invOk} warn={invOk && st.diffPct !== null && Math.abs(st.diffPct) > 3}>
          <b>{tr("3. Faktura")}</b> — {invOk ? st.invoices.map((i) => `${i.number} (${money(i.amount, i.currency)})`).join(", ") : tr("nije uvezena")}
          {!invOk && <Sub>{tr("Napravi fakturu u Vyfakturuju i uvezi PDF: Objekti → 📥 Uvoz faktura.")}</Sub>}
          {st.expectedCzk > 0 && (
            <Sub>
              {tr("Po satima × cijeni očekivano bez DPH:")}{" "}<b>{czk(st.expectedCzk)}</b>
              {invOk && <> {tr("· na fakturi bez DPH:")}{" "}<b>{czk(st.invNetCzk)}</b>
                {st.diffPct !== null && <span style={{ color: Math.abs(st.diffPct) > 3 ? S.amber : S.green, fontWeight: 700 }}> ({st.diffPct > 0 ? "+" : ""}{st.diffPct} %)</span>}</>}
            </Sub>
          )}
          {st.expectedCzk === 0 && st.hours > 0 && <Sub>{tr("Objekt nema upisanu cijenu naplate, pa ne mogu usporediti s fakturom.")}</Sub>}
          {invOk && <Sub>{tr("Plaćeno:")}{" "}{st.paid > 0 ? czk(st.paid) : tr("još nije upisano")}</Sub>}
        </Check>

        <div style={{ marginTop: 10 }}>
          <Btn onClick={() => { toggle(o.id, "done", o.name); if (!done) setIdx(idx + 1); }} kind={done ? "ghost" : "primary"} style={{ width: "100%" }}>
            {done ? tr("↺ Ponovno otvori ovaj objekt") : hoursOk && checksOk && invOk ? tr("✓ Objekt zatvoren — dalje") : tr("✓ Zatvori svejedno — dalje")}
          </Btn>
        </div>
      </>
    );
  } else if (cur.kind === "payouts") {
    const withHours = [...new Set(data.logs.filter((l) => monthKey(l.date) === mo).map((l) => l.workerId))];
    const notProposed = withHours.filter((id) => !paidFor(data, id, mo));
    const notApproved = withHours.filter((id) => paidFor(data, id, mo) && !paidFor(data, id, mo).approved);
    const ok = withHours.length > 0 && notProposed.length === 0 && notApproved.length === 0;
    body = (
      <>
        <div style={{ fontWeight: 800, fontSize: 16, marginBottom: 6 }}>{tr("💶 Isplate radnicima")}</div>
        <Check ok={notProposed.length === 0 && withHours.length > 0} warn={notProposed.length > 0}>
          <b>{tr("Predložene isplate")}</b> — {withHours.length - notProposed.length}/{withHours.length} {tr("radnika")}
          {notProposed.length > 0 && <Sub>{tr("Bez prijedloga:")}{" "}{notProposed.slice(0, 10).map(wName).join(", ")}{notProposed.length > 10 ? " …" : ""} · <Link onClick={() => goTab("obracun")}>{tr("otvori Obračun ›")}</Link></Sub>}
        </Check>
        <Check ok={notApproved.length === 0 && withHours.length > 0} warn={notApproved.length > 0}>
          <b>{tr("Odobrene isplate")}</b> — {notApproved.length ? tr("{0} čeka odobrenje", notApproved.length) : tr("sve odobreno")}
          {notApproved.length > 0 && <Sub><Link onClick={() => openAdminPanel("approvals")}>{tr("Otvori odobravanje ›")}</Link></Sub>}
        </Check>
        <Btn onClick={() => setIdx(idx + 1)} style={{ width: "100%", marginTop: 10 }}>{ok ? tr("✓ Dalje") : tr("Dalje")}</Btn>
      </>
    );
  } else if (cur.kind === "vat") {
    const v = vatSummary(data, mo);
    const dd = daysUntil(v.due);
    body = (
      <>
        <div style={{ fontWeight: 800, fontSize: 16, marginBottom: 6 }}>{tr("🧾 DPH za")}{" "}{label}</div>
        <Check ok={v.out.length > 0}>
          <b>{tr("Izdane fakture")}</b> — {v.out.length}{tr(", DPH")}{" "}{czk(v.outVat)}
        </Check>
        <Check ok={isDone(null, "purchases")}>
          <b>{tr("Primljene fakture (odbitak)")}</b> — {v.inn.length}{tr(", DPH")}{" "}{czk(v.inVat)}
          <Sub>{tr("Jesi li upisao sve primljene fakture s DPH-om za ovaj mjesec?")}{" "}<Link onClick={() => openAdminPanel("vat")}>{tr("Upiši ›")}</Link></Sub>
          <label style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 4, fontSize: 12.5, cursor: "pointer" }}>
            <input type="checkbox" checked={isDone(null, "purchases")} onChange={() => toggle(null, "purchases", "primljene fakture")} style={{ width: 16, height: 16 }} />
            {tr("Da, sve su upisane")}
          </label>
        </Check>
        <div style={{ background: v.payable > 0 ? S.amberSoft : S.greenSoft, borderRadius: 10, padding: "9px 12px", marginTop: 8, fontSize: 14 }}>
          {v.payable >= 0 ? tr("Za platiti") : tr("Vraća se vama")}: <b className="num">{czk(Math.abs(v.payable))}</b><br />
          {tr("Rok:")}{" "}<b>{fmtDate(v.due)}</b> <span style={{ color: dd < 0 ? S.red : S.sub }}>({dd >= 0 ? tr("za {0} dana", dd) : tr("kasni {0} dana", -dd)})</span>
        </div>
        <label style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 8, fontSize: 13.5, cursor: "pointer" }}>
          <input type="checkbox" checked={isDone(null, "vat_paid")} onChange={() => toggle(null, "vat_paid", "DPH plaćen")} style={{ width: 17, height: 17 }} />
          {tr("DPH prijavljen i plaćen")}
        </label>
        <Btn onClick={() => setIdx(idx + 1)} style={{ width: "100%", marginTop: 10 }}>{tr("Dalje")}</Btn>
      </>
    );
  } else {
    const doneObjs = objs.filter((o) => isDone(o.id, "done"));
    body = (
      <>
        <div style={{ fontWeight: 800, fontSize: 16, marginBottom: 6 }}>{tr("📋 Pregled —")}{" "}{label}</div>
        {objs.map((o, i) => {
          const st = objectCloseState(data, o, mo);
          return (
            <div key={o.id} onClick={() => setIdx(i)} style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", borderTop: `1px solid ${S.line}`, fontSize: 13.5, cursor: "pointer" }}>
              <span>{isDone(o.id, "done") ? "✅" : "⬜"} {o.name}</span>
              <span className="num" style={{ color: S.sub, fontSize: 12.5 }}>{fmtH(st.hours)} · {st.invoices.length ? czk(st.invNetCzk) : tr("bez fakture")}</span>
            </div>
          );
        })}
        <div style={{ borderTop: `1px solid ${S.line}`, padding: "7px 0", fontSize: 13.5 }}>{isDone(null, "vat_paid") ? "✅" : "⬜"} {tr("DPH prijavljen i plaćen")}</div>
        <div style={{ marginTop: 8, fontWeight: 700, color: doneObjs.length === objs.length ? S.green : S.amber }}>
          {doneObjs.length === objs.length ? tr("Mjesec je zatvoren za sve objekte. 🎉") : tr("Zatvoreno {0} od {1} objekata — klikni na objekt da nastaviš.", doneObjs.length, objs.length)}
        </div>
        <Btn kind="ghost" onClick={onClose} style={{ width: "100%", marginTop: 10 }}>{tr("Natrag na Danas")}</Btn>
      </>
    );
  }

  return (
    <div>
      <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA", padding: "10px 14px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <Btn small kind="ghost" onClick={onClose}>✕</Btn>
          <div style={{ textAlign: "center" }}>
            <div style={{ fontWeight: 800 }}>{tr("✅ Zatvaranje mjeseca")}</div>
            <div style={{ display: "flex", alignItems: "center", gap: 6, justifyContent: "center", fontSize: 13 }}>
              <button onClick={() => shiftMo(-1)} style={{ background: "none", border: "none", cursor: "pointer", fontWeight: 800 }}>‹</button>
              {label}
              <button onClick={() => shiftMo(1)} style={{ background: "none", border: "none", cursor: "pointer", fontWeight: 800 }}>›</button>
            </div>
          </div>
          <span className="num" style={{ fontSize: 12.5, color: S.sub, minWidth: 40, textAlign: "right" }}>{Math.min(idx, steps.length - 1) + 1}/{steps.length}</span>
        </div>
        <div style={{ display: "flex", gap: 3, marginTop: 8 }}>
          {steps.map((s, i) => {
            const done = s.kind === "object" ? isDone(s.o.id, "done") : s.kind === "vat" ? isDone(null, "vat_paid") : false;
            return <div key={i} onClick={() => setIdx(i)} title={s.kind === "object" ? s.o.name : s.kind}
              style={{ flex: 1, height: 6, borderRadius: 3, cursor: "pointer", background: i === idx ? S.blue : done ? S.green : S.line }} />;
          })}
        </div>
      </Card>
      <Card>{body}</Card>
      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
        <Btn kind="ghost" onClick={() => setIdx(Math.max(0, idx - 1))} disabled={idx === 0} style={{ flex: 1 }}>{tr("◀ Natrag")}</Btn>
        <Btn kind="ghost" onClick={() => setIdx(Math.min(steps.length - 1, idx + 1))} disabled={idx >= steps.length - 1} style={{ flex: 1 }}>{tr("Preskoči ▶")}</Btn>
      </div>
    </div>
  );
}

function WorkersTab({ data, api, onOpen, onOpenObject }) {
  const [adding, setAdding] = useState(false);
  const [showObjects, setShowObjects] = useState(true);
  const [showToday, setShowToday] = useState(false);
  const [form, setForm] = useState({ lastName: "", firstName: "", phone: "", rate: "", rateCur: "EUR", objectId: "", position: "", note: "", permitExpiry: "", contractExpiry: "" });
  const [newObj, setNewObj] = useState("");
  const [confirmObj, setConfirmObj] = useState(null);
  const [importReview, setImportReview] = useState(null);
  const [importBusy, setImportBusy] = useState(false);
  const importFileRef = useRef(null);
  const [showBulkArchive, setShowBulkArchive] = useState(false);
  const [reminderText, setReminderText] = useState("");
  const [reminderDate, setReminderDate] = useState("");
  const [archiveSel, setArchiveSel] = useState(() => new Set());
  const [archiveBusy, setArchiveBusy] = useState(false);
  const [payMonth, setPayMonth] = useState(() => { const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; });
  const [showNameFix, setShowNameFix] = useState(false);
  const [confirmSwap, setConfirmSwap] = useState(false);
  const [swapBusy, setSwapBusy] = useState(false);

  const objName = (id) => data.objects.find((o) => o.id === id)?.name || "";
  const mk = curMonth();
  const warns = expiryWarnings(data.workers, data.settings?.expiry_warn_days);
  const staleRates = rateStaleWarnings(data.workers, data.rateChanges);
  const [q, setQ] = useState("");
  const newFullName = `${form.lastName.trim()} ${form.firstName.trim()}`.trim();
  const nameDupe = newFullName && data.workers.find((w) => w.name.trim().toLowerCase() === newFullName.toLowerCase());

  const addWorker = async () => {
    if (!newFullName) return;
    if (await api.addWorker({ ...form, name: newFullName })) { setForm({ lastName: "", firstName: "", phone: "", rate: "", rateCur: "EUR", objectId: "", position: "", note: "", permitExpiry: "", contractExpiry: "" }); setAdding(false); }
  };

  const downloadWorkerTemplate = () => {
    const objects = sortedObjects(data.objects.filter((o) => !o.archived));
    downloadStyledXlsx({
      fileName: "Predlozak_radnici.xlsx",
      sheets: [
        { name: "Radnici", noTitle: true, headerColor: XC.green,
          columns: [
            { key: "ln", header: "Prezime", width: 18, bold: true }, { key: "fn", header: "Ime", width: 16, bold: true },
            { key: "phone", header: "Telefon", width: 16 }, { key: "rate", header: "Satnica", width: 10, type: "money" },
            { key: "cur", header: "Valuta satnice", width: 13, type: "center" }, { key: "obj", header: "Glavni objekt", width: 22 },
            { key: "pos", header: "Pozicija", width: 16 }, { key: "permit", header: "Istek radne dozvole", width: 18, type: "center" },
            { key: "contract", header: "Istek ugovora", width: 15, type: "center" }, { key: "note", header: "Napomena", width: 28 },
          ],
          rows: [{ ln: "Horvat", fn: "Ivan", phone: "091 234 5678", rate: 7.5, cur: "EUR", obj: objects[0]?.name || "", pos: "HSK", permit: "", contract: "", note: "" }],
          after: (ws) => {
            ws.autoFilter = undefined;
            ws.dataValidations.add("E2:E500", { type: "list", allowBlank: true, formulae: ['"EUR,CZK"'], showErrorMessage: true, errorTitle: "Valuta", error: "Odaberi EUR ili CZK." });
            if (objects.length) ws.dataValidations.add("F2:F500", { type: "list", allowBlank: true, formulae: [`'Objekti'!$A$5:$A$${4 + objects.length}`] });
            ws.getRow(2).eachCell((c) => { c.font = { ...c.font, italic: true, color: { argb: XC.muted } }; });
          } },
        { name: "Upute", title: "Kako popuniti predložak", subtitle: "Uvoz radnika u Evidenciju rada", tabColor: XC.amber,
          columns: [{ key: "a", header: "Stupac", width: 22, bold: true }, { key: "b", header: "Što upisati", width: 70, wrap: true }],
          rows: [
            { a: "Red 2 (sivo)", b: "Primjer. Prepiši ga svojim podacima ili ga obriši prije uvoza." },
            { a: "Prezime, Ime", b: "Obavezno. Jedan radnik u svakom retku." },
            { a: "Telefon", b: "Po želji, npr. +420 777 123 456." },
            { a: "Satnica", b: "Broj, npr. 150 ili 7,5." },
            { a: "Valuta satnice", b: "Odaberi s popisa: EUR ili CZK." },
            { a: "Glavni objekt", b: "Odaberi s popisa. Naziv mora biti isti kao u aplikaciji (popis je na listu „Objekti”)." },
            { a: "Pozicija", b: "Po želji, npr. pokojská, kuchyň, HSK." },
            { a: "Istek dozvole / ugovora", b: "Datum, npr. 31.12.2026. Ostavi prazno ako ne znaš." },
            { a: "Uvoz", b: "U aplikaciji: Radnici → 📥 Uvezi iz Excela, pa provjeri popis prije potvrde." },
          ] },
        { name: "Objekti", title: "Objekti u aplikaciji", subtitle: "Za stupac „Glavni objekt”", tabColor: XC.sub,
          columns: [{ key: "n", header: "Objekt", width: 30, bold: true }, { key: "c", header: "Država", width: 10, type: "center" }],
          rows: objects.map((o) => ({ n: o.name, c: o.country })) },
      ],
    }).catch((e) => alert("Izvoz nije uspio: " + (e.message || e)));
  };
  const handleImportFile = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wbk = XLSX.read(buf, { type: "array", cellDates: true });
    const sheet = wbk.Sheets[wbk.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { raw: false, dateNF: "yyyy-mm-dd", defval: "" });
    const rows = json.map((row) => {
      const lastName = String(row["Prezime"] || "").trim();
      const firstName = String(row["Ime"] || "").trim();
      // podrška i za stari predložak s jednim stupcem "Ime i prezime" (ako ga je netko već imao spremljenog)
      const name = (lastName || firstName) ? `${lastName} ${firstName}`.trim() : String(row["Ime i prezime"] || "").trim();
      const objectName = String(row["Glavni objekt"] || "").trim();
      const ob = objectName ? data.objects.find((o) => o.name.toLowerCase() === objectName.toLowerCase()) : null;
      const curRaw = String(row["Valuta satnice"] || "EUR").toUpperCase();
      return {
        name, phone: String(row["Telefon"] || "").trim(), rate: String(row["Satnica"] || ""),
        rateCur: curRaw.includes("K") ? "CZK" : "EUR",
        objectName, objectId: ob?.id || "", objectMissing: !!objectName && !ob,
        position: String(row["Pozicija"] || "").trim(),
        permitExpiry: String(row["Istek radne dozvole"] || "").trim(), contractExpiry: String(row["Istek ugovora"] || "").trim(),
        note: String(row["Napomena"] || "").trim(), include: true,
      };
    }).filter((r) => r.name);
    setImportReview(rows);
  };
  const confirmImport = async () => {
    const toImport = importReview.filter((r) => r.include);
    if (!toImport.length) return;
    setImportBusy(true);
    if (await api.addWorkersBulk(toImport)) setImportReview(null);
    setImportBusy(false);
  };
  const submitBulkArchive = async () => {
    if (!archiveSel.size) return;
    setArchiveBusy(true);
    if (await api.archiveWorkersBulk([...archiveSel])) setArchiveSel(new Set());
    setArchiveBusy(false);
  };
  const addReminderClick = async () => {
    const t = reminderText.trim();
    if (!t) return;
    if (await api.addReminder(t, reminderDate)) { setReminderText(""); setReminderDate(""); }
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
  const objectsSorted = sortedObjects(data.objects.filter((o) => !o.archived));

  // čišćenje podataka: bez imena, ili ime čiji format/velika-mala slova nisu dosljedni
  const nameIssues = data.workers.map((w) => {
    const blank = !w.name || !w.name.trim();
    const normalized = normalizeName(w.name);
    const badCase = !blank && w.name !== normalized;
    const oneWord = !blank && !looksLikeFullName(w.name);
    return (blank || badCase || oneWord) ? { w, blank, badCase, oneWord, normalized } : null;
  }).filter(Boolean);
  const fixName = (w, normalized) => api.updWorker(w.id, { ...w, rate: String(w.rate || ""), rateCur: wCur(w), name: normalized }, normalized);
  // jednokratna zamjena redoslijeda za sve koji još imaju staro "Ime Prezime" umjesto "Prezime Ime"
  const swappable = data.workers.filter((w) => looksLikeFullName(w.name));
  const doSwapAll = async () => {
    setSwapBusy(true);
    if (await api.swapAllNameOrders(swappable)) setConfirmSwap(false);
    setSwapBusy(false);
  };

  return (
    <>
      {!api.admin && (() => {
        const mine = data.logs.filter((l) => l.createdBy === api.uid() && monthKey(l.date) === mk);
        if (!mine.length) return null;
        const hrs = round2(mine.reduce((s, l) => s + l.hours, 0));
        const touched = new Set(mine.map((l) => l.workerId)).size;
        return (
          <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA" }}>
            <div style={{ fontWeight: 700, color: S.blue, marginBottom: 4 }}>{tr("👤 Tvoj unos ovaj mjesec")}</div>
            <div className="num" style={{ fontSize: 13.5 }}>{tr("Upisano")}{" "}<b>{fmtH(hrs)}</b> {tr("za")}{" "}{touched} {touched === 1 ? tr("radnika") : tr("radnika")}.</div>
          </Card>
        );
      })()}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontSize: 17, fontWeight: 700 }}>{api.t("workersTitle")} ({actives.length})</div>
        <div style={{ display: "flex", gap: 6 }}>
          <Btn small onClick={() => setAdding(!adding)}>{adding ? api.t("close") : api.t("addWorker")}</Btn>
        </div>
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
        <Btn small kind="ghost" onClick={downloadWorkerTemplate} style={{ flex: 1 }}>{tr("⬇ Predložak (Excel)")}</Btn>
        <Btn small kind="ghost" onClick={() => importFileRef.current?.click()} style={{ flex: 1 }}>{tr("📥 Uvezi iz Excela")}</Btn>
        <input ref={importFileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleImportFile} style={{ display: "none" }} />
      </div>

      {actives.length > 0 && (
        <Card style={{ background: S.amberSoft, borderColor: "#EBD9B4" }}>
          <div onClick={() => setShowBulkArchive(!showBulkArchive)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
            <div style={{ fontWeight: 700, color: S.amber }}>{tr("📁 Grupno arhiviranje (kraj sezone)")}{archiveSel.size > 0 ? ` (${archiveSel.size})` : ""}</div>
            <span style={{ color: S.amber, fontWeight: 700 }}>{showBulkArchive ? "▲" : "▼"}</span>
          </div>
          {showBulkArchive && (
            <div style={{ marginTop: 10 }}>
              <div style={{ fontSize: 12, color: S.sub, marginBottom: 8 }}>{tr("Označi radnike koji su završili s radom i klikni jednom \"Označi odabrane kao završene\".")}</div>
              <div style={{ maxHeight: 260, overflowY: "auto", marginBottom: 10 }}>
                {actives.map((w) => (
                  <label key={w.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px solid #EBD9B4", fontSize: 13.5, cursor: "pointer" }}>
                    <input type="checkbox" checked={archiveSel.has(w.id)}
                      onChange={(e) => setArchiveSel((p) => { const n = new Set(p); e.target.checked ? n.add(w.id) : n.delete(w.id); return n; })}
                      style={{ width: 17, height: 17 }} />
                    <span style={{ flex: 1 }}>{w.name}</span>
                    <span style={{ fontSize: 12, color: S.sub }}>{objName(w.objectId) || tr("bez objekta")}</span>
                  </label>
                ))}
              </div>
              <Btn small onClick={submitBulkArchive} disabled={archiveBusy || archiveSel.size === 0} style={{ width: "100%" }}>
                {archiveBusy ? tr("Spremam…") : tr("📁 Označi odabrane kao završene ({0})", archiveSel.size)}
              </Btn>
            </div>
          )}
        </Card>
      )}

      {importReview && (
        <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA" }}>
          <div style={{ fontWeight: 700, color: S.blue, marginBottom: 8 }}>{tr("📥 Pregled uvoza —")}{" "}{importReview.length} {tr("radnika u datoteci")}</div>
          {importReview.map((r, i) => (
            <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: 8, padding: "7px 0", borderBottom: "1px solid #CBDCEA" }}>
              <input type="checkbox" checked={r.include} onChange={(e) => setImportReview((p) => p.map((row, j) => j === i ? { ...row, include: e.target.checked } : row))}
                style={{ width: 18, height: 18, marginTop: 2 }} />
              <div style={{ flex: 1, fontSize: 13 }}>
                <div style={{ fontWeight: 700 }}>{r.name}</div>
                <div className="num" style={{ color: S.sub, fontSize: 12 }}>
                  {r.phone || tr("bez broja")} · {r.rate ? money(parseNum(r.rate) || 0, r.rateCur) + "/h" : tr("bez satnice")}{r.position ? " · " + r.position : ""}
                </div>
                {r.objectName && (
                  <div style={{ fontSize: 12, color: r.objectMissing ? S.red : S.sub }}>
                    {r.objectMissing ? tr("⚠️ objekt \"{0}\" ne postoji — dodaj ga prvo pa ponovi uvoz, ili će radnik ostati bez objekta", r.objectName) : "🏨 " + r.objectName}
                  </div>
                )}
              </div>
            </div>
          ))}
          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <Btn small onClick={confirmImport} disabled={importBusy || !importReview.some((r) => r.include)}>
              {importBusy ? tr("Uvozim…") : tr("✓ Uvezi ({0})", importReview.filter((r) => r.include).length)}
            </Btn>
            <Btn small kind="ghost" onClick={() => setImportReview(null)}>{tr("Odustani")}</Btn>
          </div>
        </Card>
      )}

      {adding && (
        <Card>
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}>
              <Field label={api.t("lastName")}><input value={form.lastName} onChange={(e) => setForm({ ...form, lastName: e.target.value })} placeholder={tr("npr. Horvat")} /></Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label={api.t("firstName")}><input value={form.firstName} onChange={(e) => setForm({ ...form, firstName: e.target.value })} placeholder={tr("npr. Ivan")} /></Field>
            </div>
          </div>
          {nameDupe && (
            <div style={{ fontSize: 12.5, color: S.amber, fontWeight: 600, margin: "-6px 0 10px" }}>
              {tr("⚠️ Radnik s ovim imenom već postoji")}{nameDupe.archived ? tr(" (u arhivi)") : ""} {tr("— provjeri da ne dupliraš.")}
            </div>
          )}
          <Field label={api.t("phone")}><input value={form.phone} inputMode="tel" onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
          <Field label={tr("Satnica ({0} / sat) — nije obavezno", form.rateCur === "CZK" ? "Kč" : "€")}>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <input value={form.rate} inputMode="decimal" onChange={(e) => setForm({ ...form, rate: e.target.value })} placeholder={form.rateCur === "CZK" ? tr("npr. 150") : tr("npr. 6.50")} />
              <CurChips small value={form.rateCur} onChange={(v) => setForm({ ...form, rateCur: v })} />
            </div>
          </Field>
          {isRateSuspicious(form.rate, form.rateCur) && (
            <div style={{ fontSize: 12.5, color: S.amber, fontWeight: 600, margin: "-6px 0 10px" }}>
              {tr("⚠️ Ova satnica izgleda neobično — provjeri nisi li se utipkala (npr. višak/manjak nule).")}
            </div>
          )}
          <Field label={tr("Glavni objekt")}><ObjectSelect data={data} api={api} value={form.objectId} onChange={(v) => {
            const ob = data.objects.find((o) => o.id === v);
            setForm({ ...form, objectId: v, rateCur: ob ? countryCur(ob.country) : form.rateCur });
          }} /></Field>
          <Field label={tr("Pozicija (npr. HSK, kuhinja, bar…) — nije obavezno")}>
            <input value={form.position} onChange={(e) => setForm({ ...form, position: e.target.value })} placeholder={tr("npr. HSK")} />
          </Field>
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}><Field label={tr("Istek radne dozvole")}><input type="date" value={form.permitExpiry} onChange={(e) => setForm({ ...form, permitExpiry: e.target.value })} /></Field></div>
            <div style={{ flex: 1 }}><Field label={tr("Istek ugovora")}><input type="date" value={form.contractExpiry} onChange={(e) => setForm({ ...form, contractExpiry: e.target.value })} /></Field></div>
          </div>
          {api.admin && <Field label={tr("Broj bankovnog računa (Číslo účtu) — za knjigovođu")}><input value={form.bankAccount || ""} onChange={(e) => setForm({ ...form, bankAccount: e.target.value })} placeholder={tr("npr. 123456789/0800")} /></Field>}
          <Field label={api.t("note") + " (OIB…)"}><textarea rows={2} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
          <Btn onClick={addWorker} style={{ width: "100%" }}>{api.t("saveWorker")}</Btn>
        </Card>
      )}

      {actives.length > 5 && (
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder={tr("🔍 Traži po imenu, telefonu, poziciji…")} style={{ marginBottom: 10 }} />
      )}

      {actives.length === 0 && !adding && <Empty text={api.t("noWorkersYet")} />}

      {actives.filter((w) => {
        const t = q.trim().toLowerCase();
        return !t || w.name.toLowerCase().includes(t) || (w.phone || "").includes(t) || (w.position || "").toLowerCase().includes(t);
      }).map((w) => {
        const h = data.logs.filter((l) => l.workerId === w.id && monthKey(l.date) === mk).reduce((s, l) => s + l.hours, 0);
        return (
          <Card key={w.id} style={{ cursor: "pointer" }}>
            <div onClick={() => onOpen(w.id)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 15.5 }}>{w.name}{w.position && <span style={{ marginLeft: 6, fontWeight: 600, fontSize: 12, color: S.blue }}>· {w.position}</span>}</div>
                <div style={{ fontSize: 13, color: S.sub, marginTop: 2 }}>{objName(w.objectId) || "—"} · {w.phone || tr("bez broja")}</div>
              </div>
              <div style={{ textAlign: "right" }}>
                {rateNow(data, w) > 0
                  ? <div className="num" style={{ fontWeight: 700, color: S.green }}>{money(rateNow(data, w), wCur(w))}/h</div>
                  : <div style={{ fontWeight: 700, color: S.amber, fontSize: 12.5 }}>{api.t("withoutRate")}</div>}
                <div className="num" style={{ fontSize: 12.5, color: S.sub }}>{fmtH(round2(h))} {api.t("thisMonth")}</div>
                {rateNow(data, w) > 0 && h > 0 && (
                  <div className="num" style={{ fontSize: 12.5, fontWeight: 700, color: S.ink, marginTop: 1 }}>{money(round2(h * rateNow(data, w)), wCur(w))} {tr("za isplatu")}</div>
                )}
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
                  <div style={{ fontSize: 13, color: S.sub, marginTop: 2 }}>{w.phone || tr("bez broja")}{w.archivedDate ? " · završio " + fmtDate(w.archivedDate) : ""}</div>
                </div>
                <Tag color={S.sub} bg="#EEF0ED">{tr("arhiva")}</Tag>
              </div>
            </Card>
          ))}
        </>
      )}

      <div style={{ fontSize: 13, fontWeight: 700, color: S.sub, textTransform: "uppercase", letterSpacing: 1, margin: "22px 0 8px" }}>{tr("Alati i provjere")}</div>
      {api.admin && staleRates.length > 0 && (
        <Card style={{ background: S.amberSoft, borderColor: "#EBD9B4" }}>
          <div style={{ fontWeight: 700, color: S.amber, marginBottom: 6 }}>{tr("💰 Bez promjene satnice godinu+ dana")}</div>
          {staleRates.map((x, i) => (
            <div key={i} onClick={() => onOpen(x.w.id)} style={{ fontSize: 13.5, padding: "4px 0", cursor: "pointer" }}>
              <b>{x.w.name}</b> {tr("— ista satnica od")}{" "}<span className="num">{fmtDate(x.since)}</span>
            </div>
          ))}
        </Card>
      )}

      {api.admin && !data.settings.name_order_migrated && swappable.length > 0 && (
        <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA" }}>
          <div style={{ fontWeight: 700, color: S.blue, marginBottom: 6 }}>{tr("🔁 Prebaci redoslijed: Ime Prezime → Prezime Ime")}</div>
          <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 8 }}>
            {tr("Svi radnici su upisani kao \"Ime Prezime\". Ovo jednom zamijeni redoslijed svima u \"Prezime Ime\" — nakon toga su popisi posloženi po abecedi prezimena. Primjer:")}{" "}{swappable[0] && <>"{swappable[0].name}" → "<b>{swapNameOrder(swappable[0].name)}</b>"</>}.
          </div>
          {!confirmSwap ? (
            <Btn small onClick={() => setConfirmSwap(true)}>{tr("Zamijeni svima (")}{swappable.length})</Btn>
          ) : (
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <Btn small onClick={doSwapAll} disabled={swapBusy}>{swapBusy ? tr("Zamjenjujem…") : tr("Potvrdi zamjenu ({0})", swappable.length)}</Btn>
              <Btn small kind="ghost" onClick={() => setConfirmSwap(false)} disabled={swapBusy}>{tr("Odustani")}</Btn>
            </div>
          )}
        </Card>
      )}

      {api.admin && nameIssues.length > 0 && (
        <Card style={{ background: S.amberSoft, borderColor: "#EBD9B4" }}>
          <div onClick={() => setShowNameFix(!showNameFix)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
            <div style={{ fontWeight: 700, color: S.amber }}>{tr("🧹 Provjeri imena (")}{nameIssues.length})</div>
            <span style={{ color: S.amber, fontWeight: 700 }}>{showNameFix ? "▲" : "▼"}</span>
          </div>
          {showNameFix && (
            <div style={{ marginTop: 8, borderTop: "1px dashed #EBD9B4", paddingTop: 8 }}>
              {nameIssues.map(({ w, blank, badCase, normalized }) => (
                <div key={w.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px dashed #EBD9B4" }}>
                  <div style={{ flex: 1, fontSize: 13.5 }}>
                    {blank ? (
                      <span onClick={() => onOpen(w.id)} style={{ color: S.red, fontWeight: 700, cursor: "pointer" }}>{tr("🚫 (bez imena) —")}{" "}{objName(w.objectId) || tr("bez objekta")}</span>
                    ) : badCase ? (
                      <span onClick={() => onOpen(w.id)} style={{ cursor: "pointer" }}>"{w.name}" → <b>"{normalized}"</b></span>
                    ) : (
                      <span onClick={() => onOpen(w.id)} style={{ cursor: "pointer" }}>"{w.name}{tr("\" — samo jedno ime, nedostaje prezime")}</span>
                    )}
                  </div>
                  {badCase && <Btn small onClick={() => fixName(w, normalized)}>{tr("Ispravi")}</Btn>}
                </div>
              ))}
              <div style={{ fontSize: 12, color: S.sub, marginTop: 8 }}>
                {tr("Za upis bez imena klikni na njega da otvoriš radnika, upiši ispravno ime i spremi (ili ga obriši ako je greškom nastao).")}
              </div>
            </div>
          )}
        </Card>
      )}

      <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA" }}>
        <div onClick={() => setShowObjects(!showObjects)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
          <div style={{ fontWeight: 700, color: S.blue }}>{tr("🏨 Objekti (")}{objectsSorted.length})</div>
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
                      {fmtH(round2(h))} {tr("ovaj mjesec")}
                      {api.admin && o.billRate ? <span style={{ color: S.amber, fontWeight: 700 }}> {tr("· naplata")}{" "}{money(o.billRate, o.billCur)}/h</span> : ""}
                    </div>
                  </div>
                  {api.admin && (
                    <button onClick={() => delObject(o)} style={{
                      background: confirmObj === o.id ? S.red : "none", color: confirmObj === o.id ? "#fff" : S.red,
                      border: "none", borderRadius: 8, cursor: "pointer",
                      fontSize: confirmObj === o.id ? 12.5 : 15, fontWeight: 700, padding: confirmObj === o.id ? "5px 10px" : 4 }}>
                      {confirmObj === o.id ? tr("Potvrdi brisanje") : "✕"}
                    </button>
                  )}
                </div>
              );
            })}
            <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
              <input value={newObj} onChange={(e) => setNewObj(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addObject()} placeholder={tr("Naziv novog objekta…")} />
              <Btn small onClick={addObject}>{tr("Dodaj")}</Btn>
            </div>
          </div>
        )}
      </Card>

      {(() => {
        const today = todayISO();
        const todayLogs = data.logs.filter((l) => l.date === today);
        const todayHours = round2(todayLogs.reduce((s, l) => s + l.hours, 0));
        const workedIds = new Set(todayLogs.map((l) => l.workerId));
        const notLoggedToday = actives.filter((w) => !workedIds.has(w.id));
        return (
          <Card style={{ background: S.greenSoft, borderColor: "#C5DED2" }}>
            <div onClick={() => setShowToday(!showToday)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
              <div style={{ fontWeight: 700, color: S.green }}>{tr("📊 Danas —")}{" "}{fmtDate(today)}</div>
              <span style={{ color: S.green, fontWeight: 700 }}>{showToday ? "▲" : "▼"}</span>
            </div>
            <div className="num" style={{ fontSize: 13.5, marginTop: 6 }}>
              <b>{fmtH(todayHours)}</b> {tr("upisano danas ·")}{" "}{workedIds.size}/{actives.length} {tr("radnika upisalo sate")}
            </div>
            {showToday && (
              <div style={{ marginTop: 8, borderTop: "1px dashed #C5DED2", paddingTop: 8 }}>
                {notLoggedToday.length === 0 ? (
                  <div style={{ fontSize: 12.5, color: S.sub }}>{tr("Svi aktivni radnici imaju upis za danas. 🎉")}</div>
                ) : (
                  <>
                    <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 4 }}>{tr("Bez upisa danas (")}{notLoggedToday.length}{tr(") — nije nužno problem, možda nisu radili:")}</div>
                    {notLoggedToday.map((w) => (
                      <span key={w.id} onClick={() => onOpen(w.id)} style={{
                        display: "inline-block", fontSize: 12.5, color: S.sub, background: "#fff", border: "1px solid #C5DED2",
                        borderRadius: 999, padding: "3px 9px", margin: "2px 4px 2px 0", cursor: "pointer" }}>{w.name}</span>
                    ))}
                  </>
                )}
                {typeof Notification !== "undefined" && Notification.permission === "default" && (
                  <Btn small kind="ghost" onClick={() => Notification.requestPermission()} style={{ marginTop: 10 }}>
                    {tr("🔔 Uključi obavijesti za istek dokumenata")}
                  </Btn>
                )}
              </div>
            )}
          </Card>
        );
      })()}

    </>
  );
}

/* ================================================================== */
/*  OBJEKTI — pregled po objektu                                       */
/* ================================================================== */
function ObjectsTab({ data, api, onOpenObject }) {
  const mk = curMonth();
  const [newObj, setNewObj] = useState("");
  // aktivni filter/tab — i zadana država za novi objekt; otvori na državi koja ima otvorene objekte
  const [country, setCountry] = useState(() => {
    const open = data.objects.filter((o) => !o.archived);
    const cz = open.filter((o) => o.country === "CZ").length, hr = open.filter((o) => (o.country || "HR") === "HR").length;
    return cz > hr ? "CZ" : "HR";
  });
  const [q, setQ] = useState("");
  const [invReview, setInvReview] = useState(null); // [{ id, file, fileName, parsed:{...}, objectId, remember, include, error }]
  const [invShowRaw, setInvShowRaw] = useState(() => new Set());
  const [invBusy, setInvBusy] = useState(false);
  const [invOcrStatus, setInvOcrStatus] = useState(""); // "" ili npr. "OCR: ime.pdf (stranica 1/2)…"
  const invImportRef = useRef(null);
  const addObject = async () => {
    const n = newObj.trim(); if (!n) return;
    if (data.objects.find((o) => o.name.toLowerCase() === n.toLowerCase())) { setNewObj(""); return; }
    await api.addObject(n, country); setNewObj("");
  };

  const matchObjectByCustomer = (customerName) => {
    if (!customerName) return "";
    const norm = (s) => (s || "").toLowerCase().trim();
    const target = norm(customerName);
    const byMap = data.objects.find((o) => o.invoiceCustomerName && norm(o.invoiceCustomerName) === target);
    if (byMap) return byMap.id;
    const byName = data.objects.find((o) => norm(o.name) === target);
    return byName ? byName.id : "";
  };
  const handleInvoiceFiles = async (e) => {
    const files = Array.from(e.target.files || []);
    e.target.value = "";
    if (!files.length) return;
    const rows = [];
    for (const file of files) {
      try {
        let pages = await extractPdfPages(file);
        let usedOcr = false;
        if (!pages.some((p) => p.length > 0)) {
          // nema teksta u PDF-u (vjerojatno skenirano/fotografirano) — pokušaj OCR (čitanje sa slike)
          usedOcr = true;
          setInvOcrStatus(tr("🔍 OCR čita: {0}…", file.name));
          pages = await ocrPdfPages(file, (p, total) => setInvOcrStatus(tr("🔍 OCR čita: {0} (stranica {1}/{2})…", file.name, p, total)));
          setInvOcrStatus("");
        }
        const invoiceChunks = splitPdfInvoices(pages);
        invoiceChunks.forEach((items, idx) => {
          const id = Date.now() + "_" + Math.random().toString(36).slice(2) + "_" + idx;
          const parsed = parseInvoicePdf(items);
          const label = (invoiceChunks.length > 1 ? `${file.name} — faktura ${idx + 1}/${invoiceChunks.length}` : file.name) + (usedOcr ? " · OCR" : "");
          rows.push({ id, file, fileName: label, parsed, objectId: matchObjectByCustomer(parsed.customerName), remember: true, include: true, ocr: usedOcr });
        });
      } catch {
        setInvOcrStatus("");
        const id = Date.now() + "_" + Math.random().toString(36).slice(2);
        rows.push({ id, file, fileName: file.name,
          parsed: { number: "", issueDate: todayISO(), amount: 0, vatAmount: 0, period: curMonth(), customerName: "" },
          objectId: "", remember: true, include: true, error: "Nisam uspio pročitati ovaj PDF — provjeri/upiši ručno." });
      }
    }
    setInvReview((prev) => (prev ? [...prev, ...rows] : rows));
  };
  const updInvRow = (id, patch) => setInvReview((rows) => rows.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  const updInvField = (id, field, value) => setInvReview((rows) => rows.map((r) => (r.id === id ? { ...r, parsed: { ...r.parsed, [field]: value } } : r)));
  const removeInvRow = (id) => setInvReview((rows) => { const n = rows.filter((r) => r.id !== id); return n.length ? n : null; });
  const submitInvImport = async () => {
    const toImport = invReview.filter((r) => r.include && r.objectId);
    if (!toImport.length) return;
    setInvBusy(true);
    const invs = [];
    const rememberPairs = [];
    for (const r of toImport) {
      const object = data.objects.find((o) => o.id === r.objectId);
      if (!object) continue;
      let filePath = "";
      const safeName = r.fileName.replace(/[^\w.\-]+/g, "_").slice(0, 80);
      const path = `${object.id}/invoice_import_${Date.now()}_${safeName}`;
      const { error: upErr } = await supabase.storage.from("invoices").upload(path, r.file);
      if (!upErr) filePath = path;
      const amount = round2(parseNum(r.parsed.amount) || 0);
      const vatAmount = round2(parseNum(r.parsed.vatAmount) || 0);
      const issueDate = r.parsed.issueDate || todayISO();
      const period = r.parsed.period || issueDate.slice(0, 7);
      const number = (r.parsed.number || "").trim() || `UVOZ-${object.name.slice(0, 3).toUpperCase()}-${period}`;
      invs.push({ objectId: object.id, number, period, issueDate, duzp: r.parsed.duzp || null, hours: 0, rate: 0, amount, currency: "CZK", note: "Uvezeno iz PDF-a", vatAmount, vatCurrency: "CZK", filePath });
      if (r.remember && r.parsed.customerName && object.invoiceCustomerName !== r.parsed.customerName) {
        rememberPairs.push({ object, name: r.parsed.customerName });
      }
    }
    if (invs.length) await api.importObjectInvoices(invs);
    for (const { object, name } of rememberPairs) await api.setInvoiceCustomerName(object, name);
    setInvBusy(false);
    setInvReview(null);
  };
  const list = sortedObjects(data.objects.filter((o) => !o.archived));
  const closedList = sortedObjects(data.objects.filter((o) => o.archived));
  const countHR = list.filter((o) => (o.country || "HR") === "HR").length;
  const countCZ = list.filter((o) => o.country === "CZ").length;
  const shown = list.filter((o) => (o.country || "HR") === country);
  const renderObject = (o) => {
        const logs = data.logs.filter((l) => l.objectId === o.id && monthKey(l.date) === mk);
        const brigadeM = (data.brigadeCosts || []).filter((b) => b.objectId === o.id && monthKey(b.date) === mk);
        const hrs = round2(logs.reduce((s, l) => s + l.hours, 0) + brigadeM.reduce((s, b) => s + b.hours, 0));
        const workerIds = new Set([
          ...data.workers.filter((w) => !w.archived && w.objectId === o.id).map((w) => w.id),
          ...logs.map((l) => l.workerId),
        ]);
        const rv = objRevenue(data, o, logs, brigadeM);
        return (
          <Card key={o.id} style={{ cursor: "pointer" }} onClick={() => onOpenObject(o.id)}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 15.5 }}>{o.name} <span style={{ color: S.blue }}>›</span></div>
                <div className="num" style={{ fontSize: 13, color: S.sub, marginTop: 2 }}>
                  {workerIds.size} {tr("radnika ·")}{" "}{fmtH(hrs)} {tr("ovaj mjesec")}
                </div>
              </div>
              {api.admin && (rv.revenue > 0 || rv.revenueKc > 0) && (
                <div className="num" style={{ textAlign: "right" }}>
                  <div style={{ fontWeight: 700, color: S.amber }}>{rv.revenue > 0 ? eur(rv.revenue) : czk(rv.revenueKc)}</div>
                  {rv.revenue > 0 && rv.revenueKc > 0 && <div style={{ fontWeight: 700, color: S.amber, fontSize: 12.5 }}>{czk(rv.revenueKc)}</div>}
                  <div style={{ fontSize: 11.5, color: S.sub }}>{tr("naplata ovaj mj.")}</div>
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

      {api.admin && (
        <>
          <Btn small kind="ghost" onClick={() => invImportRef.current?.click()} disabled={!!invOcrStatus} style={{ width: "100%", marginBottom: invOcrStatus ? 4 : 14 }}>
            {tr("📥 Uvoz faktura (PDF, više odjednom)")}
          </Btn>
          {invOcrStatus && (
            <div style={{ fontSize: 12.5, color: S.blue, fontWeight: 600, marginBottom: 14 }}>
              {invOcrStatus} {tr("(sken bez teksta — ovo može potrajati pola minute po stranici)")}
            </div>
          )}
          <input ref={invImportRef} type="file" accept="application/pdf" multiple onChange={handleInvoiceFiles} style={{ display: "none" }} />
        </>
      )}

      {invReview && (
        <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA", marginBottom: 14 }}>
          <div style={{ fontWeight: 700, color: S.blue, marginBottom: 8 }}>{tr("📥 Pregled uvezenih faktura (")}{invReview.length})</div>
          {invReview.map((r) => {
            const matched = !!r.objectId;
            return (
              <div key={r.id} style={{ background: "#fff", borderRadius: 10, padding: "10px 12px", marginBottom: 8 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                  <input type="checkbox" checked={r.include} onChange={(e) => updInvRow(r.id, { include: e.target.checked })} style={{ width: 18, height: 18 }} />
                  <span style={{ flex: 1, fontSize: 12.5, fontWeight: 700, color: S.sub }}>{r.fileName}</span>
                  <button onClick={() => removeInvRow(r.id)} style={{ background: "none", border: "none", color: S.red, fontSize: 15, cursor: "pointer", padding: 2 }}>✕</button>
                </div>
                {r.error && <div style={{ fontSize: 12, color: S.red, marginBottom: 6 }}>⚠️ {r.error}</div>}
                {r.ocr && !r.error && (
                  <div style={{ fontSize: 11.5, color: S.amber, fontWeight: 600, marginBottom: 6 }}>
                    {tr("🔍 Pročitano OCR-om sa skena — manje pouzdano, pažljivo provjeri brojeve ispod prije uvoza.")}
                  </div>
                )}
                <div style={{ marginBottom: 6 }}>
                  <select value={r.objectId} onChange={(e) => updInvRow(r.id, { objectId: e.target.value })}
                    style={{ borderColor: matched ? S.green : S.amber, fontWeight: 700, color: matched ? S.green : S.amber }}>
                    <option value="">{tr("— odaberi objekt —")}</option>
                    {sortedObjects(data.objects.filter((o) => !o.archived)).map((o) => <option key={o.id} value={o.id}>{o.name}</option>)}
                  </select>
                  {r.parsed.customerName && (
                    <div style={{ fontSize: 11.5, color: S.sub, marginTop: 3 }}>
                      {tr("Kupac na fakturi: \"")}{r.parsed.customerName}"{matched ? tr(" — prepoznato ✓") : tr(" — nije prepoznato, odaberi ručno")}
                    </div>
                  )}
                </div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
                  <input value={r.parsed.number} onChange={(e) => updInvField(r.id, "number", e.target.value)} placeholder={tr("broj fakture")} style={{ flex: "1 1 110px" }} />
                  <input type="date" value={r.parsed.issueDate} onChange={(e) => updInvField(r.id, "issueDate", e.target.value)} title={tr("Datum vystavení")} style={{ flex: "1 1 130px" }} />
                  <label style={{ flex: "1 1 130px", fontSize: 11, color: S.sub }}>{tr("DUZP (mjesec za DPH)")}
                    <input type="date" value={r.parsed.duzp || ""} onChange={(e) => updInvField(r.id, "duzp", e.target.value)} />
                  </label>
                  <input inputMode="decimal" value={r.parsed.amount} onChange={(e) => updInvField(r.id, "amount", e.target.value)} placeholder={tr("iznos Kč")} style={{ flex: "1 1 100px" }} />
                  <input inputMode="decimal" value={r.parsed.vatAmount} onChange={(e) => updInvField(r.id, "vatAmount", e.target.value)} placeholder={tr("PDV Kč")} style={{ flex: "1 1 100px" }} />
                </div>
                {matched && r.parsed.customerName && (
                  <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: S.sub, cursor: "pointer" }}>
                    <input type="checkbox" checked={r.remember} onChange={(e) => updInvRow(r.id, { remember: e.target.checked })} style={{ width: 15, height: 15 }} />
                    {tr("Zapamti ovaj naziv za ubuduće automatsko prepoznavanje")}
                  </label>
                )}
                {!matched && r.parsed.customerName && (
                  <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: S.sub, cursor: "pointer" }}>
                    <input type="checkbox" checked={r.remember} onChange={(e) => updInvRow(r.id, { remember: e.target.checked })} style={{ width: 15, height: 15 }} />
                    {tr("Zapamti \"")}{r.parsed.customerName}{tr("\" za odabrani objekt (kad ga odabereš)")}
                  </label>
                )}
                <button onClick={() => setInvShowRaw((s) => { const n = new Set(s); n.has(r.id) ? n.delete(r.id) : n.add(r.id); return n; })}
                  style={{ background: "none", border: "none", color: S.blue, fontSize: 11.5, cursor: "pointer", padding: "4px 0", marginTop: 4 }}>
                  {invShowRaw.has(r.id) ? tr("▲ Sakrij pročitani tekst") : tr("▼ Prikaži pročitani tekst iz PDF-a (ako iznos/PDV nisu točni)")}
                </button>
                {invShowRaw.has(r.id) && (
                  <div style={{ fontSize: 10.5, fontFamily: "monospace", background: S.bg, borderRadius: 8, padding: 8, marginTop: 4, maxHeight: 180, overflowY: "auto", whiteSpace: "pre-wrap", color: S.sub }}>
                    {r.parsed.rawText || tr("(ništa pročitano)")}
                  </div>
                )}
              </div>
            );
          })}
          <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
            <Btn onClick={submitInvImport} disabled={invBusy || !invReview.some((r) => r.include && r.objectId)} style={{ flex: 1 }}>
              {invBusy ? tr("Uvozim…") : tr("✓ Uvezi ({0})", invReview.filter((r) => r.include && r.objectId).length)}
            </Btn>
            <Btn kind="ghost" onClick={() => setInvReview(null)} disabled={invBusy}>{tr("Odustani")}</Btn>
          </div>
          <div style={{ fontSize: 11.5, color: S.sub, marginTop: 8 }}>
            {tr("Podaci se čitaju automatski iz PDF-a (broj, datum, iznos, PDV, kupac) — provjeri prije uvoza, po potrebi ispravi ručno.")}
          </div>
        </Card>
      )}

      {shown.length > 5 && (
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder={tr("🔍 Traži objekt po imenu…")} style={{ marginBottom: 12 }} />
      )}
      {shown.length === 0 ? <Empty text={tr("Još nema objekata u: {0}.", COUNTRY_NAME[country])} /> : (() => {
        const filtered = shown.filter((o) => o.name.toLowerCase().includes(q.trim().toLowerCase()));
        return filtered.length === 0 ? <Empty text={tr("Nema rezultata.")} /> : filtered.map(renderObject);
      })()}

      {closedList.length > 0 && (
        <>
          <div style={{ fontSize: 15, fontWeight: 700, color: S.sub, margin: "18px 0 8px" }}>{tr("📁 Zatvoreni objekti (")}{closedList.length})</div>
          {closedList.map((o) => (
            <Card key={o.id} style={{ opacity: 0.8 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div onClick={() => onOpenObject(o.id)} style={{ cursor: "pointer" }}>
                  <div style={{ fontWeight: 700, fontSize: 15 }}>{FLAG[o.country || "HR"]} {o.name}</div>
                  <div style={{ fontSize: 12.5, color: S.sub, marginTop: 2 }}>{o.archivedDate ? "zatvoren " + fmtDate(o.archivedDate) : ""}</div>
                </div>
                {api.admin && <Btn small kind="ghost" onClick={() => api.archiveObject(o, false)}>{tr("↩ Otvori")}</Btn>}
              </div>
            </Card>
          ))}
        </>
      )}
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
  const objName = (id) => data.objects.find((o) => o.id === id)?.name || "";
  const exportDirectory = () => {
    const list = [...data.workers].sort((a, b) => (a.archived - b.archived) || a.name.localeCompare(b.name, "hr"));
    const soon = (iso) => { if (!iso) return false; const d = (new Date(iso) - new Date()) / 86400000; return d < (Number(api.settings.expiry_warn_days) || 30); };
    downloadStyledXlsx({
      fileName: `Imenik_radnika_${todayISO()}.xlsx`,
      sheets: [{
        name: "Imenik", title: "Imenik radnika", subtitle: `${active} aktivnih · ${former} bivših`, landscape: true, freezeCols: 1,
        columns: [
          { key: "name", header: "Prezime i ime", width: 26, bold: true }, { key: "phone", header: "Telefon", width: 16 }, { key: "acc", header: "Broj računa", width: 22 },
          { key: "obj", header: "Objekt", width: 20 }, { key: "pos", header: "Pozicija", width: 18 },
          { key: "rate", header: "Satnica", width: 10, type: "money" }, { key: "cur", header: "Valuta", width: 8, type: "center" },
          { key: "permit", header: "Istek radne dozvole", width: 14, type: "date", color: (v, r) => (r.permitSoon ? XC.red : null), fill: (v, r) => (r.permitSoon ? XC.redSoft : null) },
          { key: "contract", header: "Istek ugovora", width: 14, type: "date", color: (v, r) => (r.contractSoon ? XC.red : null), fill: (v, r) => (r.contractSoon ? XC.redSoft : null) },
          { key: "status", header: "Status", width: 18, type: "center", color: (v) => (v === "Aktivan" ? XC.green : XC.muted) },
          { key: "note", header: "Napomena", width: 32, wrap: true },
        ],
        rows: list.map((w) => ({
          name: w.name, phone: w.phone || "", acc: w.bankAccount || "", obj: objName(w.objectId), pos: w.position || "",
          rate: rateNow(data, w) || "", cur: wCur(w) === "CZK" ? "Kč" : "€",
          permit: w.permitExpiry, contract: w.contractExpiry,
          permitSoon: !w.archived && soon(w.permitExpiry), contractSoon: !w.archived && soon(w.contractExpiry),
          status: w.archived ? "Bivši" + (w.archivedDate ? " (" + fmtDate(w.archivedDate) + ")" : "") : "Aktivan", note: w.note || "",
        })),
        note: "Crveno označeni datumi ističu uskoro ili su već istekli.",
      }],
    }).catch((e) => alert("Izvoz nije uspio: " + (e.message || e)));
  };
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
          <span style={{ color: S.green, fontWeight: 700 }}>{active} {tr("aktivnih")}</span>{former > 0 && <> · {former} {tr("bivših")}</>}
        </div>
      </div>
      {api.admin && <Btn small kind="excel" onClick={exportDirectory} style={{ width: "100%", marginBottom: 10 }}>{tr("📊 Izvezi cijeli imenik (Excel)")}</Btn>}
      <input value={q} onChange={(e) => setQ(e.target.value)} placeholder={api.t("searchPh")} style={{ marginBottom: 12 }} />
      {list.length === 0 ? <Empty text={tr("Nema rezultata.")} /> : groups.map((g) => (
        <div key={g.letter}>
          <div style={{ fontSize: 13, fontWeight: 800, color: S.blue, padding: "6px 4px 4px" }}>{g.letter}</div>
          <Card style={{ padding: "4px 14px" }}>
            {g.workers.map((w, i) => (
              <div key={w.id} onClick={() => onOpen(w.id)} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 0", cursor: "pointer",
                borderBottom: i < g.workers.length - 1 ? `1px solid ${S.line}` : "none", opacity: w.archived ? 0.75 : 1 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 15 }}>{w.name}{w.position && <span style={{ marginLeft: 6, fontWeight: 600, fontSize: 12, color: S.blue }}>· {w.position}</span>}</div>
                  <div style={{ fontSize: 12.5, color: S.sub, marginTop: 1 }}>
                    {w.phone ? <a href={"tel:" + w.phone} onClick={(e) => e.stopPropagation()} style={{ color: S.sub }}>📞 {w.phone}</a> : tr("bez broja")}
                  </div>
                </div>
                {w.archived
                  ? <Tag color={S.sub} bg="#EEF0ED">{tr("Završio")}{w.archivedDate ? " " + fmtDate(w.archivedDate) : ""}</Tag>
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
/*  DETALJ ZAPOSLENIKA (provizija po mjesecima, detaljno po objektu)   */
/* ================================================================== */
function ProfileDetail({ profile, data, api, onBack }) {
  const [payForm, setPayForm] = useState({ amount: "", cur: "EUR", date: todayISO(), note: "" });
  const monthLabel = (mk) => `${mName(Number(mk.slice(5, 7)) - 1)} ${mk.slice(0, 4)}.`;
  const myObjIds = [...new Set((data.objectMembers || []).filter((m) => m.user_id === profile.id).map((m) => m.object_id))];
  const myObjects = myObjIds.map((oid) => data.objects.find((o) => o.id === oid)).filter(Boolean).sort((a, b) => a.name.localeCompare(b.name, "hr"));
  const rateFor = (objId) => {
    const ov = (data.commissionRates || []).find((c) => c.userId === profile.id && c.objectId === objId);
    return ov ? { rate: ov.rate, cur: ov.currency } : { rate: Number(profile.hourly_commission) || 0, cur: profile.commission_currency || "EUR" };
  };
  const logsForMe = data.logs.filter((l) => myObjIds.includes(l.objectId));
  const months = [...new Set(logsForMe.map((l) => monthKey(l.date)))].sort().reverse();

  const monthRows = months.map((mk) => {
    const perObj = myObjects.map((o) => {
      const hrs = round2(logsForMe.filter((l) => l.objectId === o.id && monthKey(l.date) === mk).reduce((s, l) => s + l.hours, 0));
      const { rate, cur } = rateFor(o.id);
      return { obj: o, hrs, rate, cur, amount: round2(hrs * rate) };
    }).filter((r) => r.hrs > 0);
    const totalsByCur = {};
    perObj.forEach((r) => { totalsByCur[r.cur] = round2((totalsByCur[r.cur] || 0) + r.amount); });
    return { mk, perObj, totalsByCur };
  }).filter((m) => m.perObj.length > 0);

  const grandTotals = {};
  monthRows.forEach((m) => Object.entries(m.totalsByCur).forEach(([c, v]) => { grandTotals[c] = round2((grandTotals[c] || 0) + v); }));

  const myPayouts = (data.commissionPayouts || []).filter((c) => c.userId === profile.id).sort((a, b) => b.payDate.localeCompare(a.payDate));
  const paidTotals = {};
  myPayouts.forEach((c) => { paidTotals[c.currency] = round2((paidTotals[c.currency] || 0) + c.amount); });

  const owedTotals = {};
  new Set([...Object.keys(grandTotals), ...Object.keys(paidTotals)]).forEach((c) => {
    owedTotals[c] = round2((grandTotals[c] || 0) - (paidTotals[c] || 0));
  });

  const savePayout = () => {
    const amt = round2(parseNum(payForm.amount) || 0);
    if (amt <= 0) return;
    api.addCommissionPayout(profile, amt, payForm.cur, payForm.date, payForm.note);
    setPayForm({ amount: "", cur: payForm.cur, date: payForm.date, note: "" });
  };

  return (
    <>
      <button onClick={onBack} style={{ background: "none", border: "none", color: S.green, fontWeight: 700, fontSize: 14.5, padding: "2px 0 12px", cursor: "pointer" }}>{tr("← Natrag")}</button>

      <Card>
        <div style={{ fontSize: 19, fontWeight: 800 }}>{profile.role === "admin" ? "👑 " : ""}{profile.name || tr("(bez imena)")}</div>
        <div style={{ color: S.sub, fontSize: 13.5, marginTop: 3 }}>
          {profile.role === "admin" ? tr("Admin") : tr("Zaposlenik")}{myObjects.length > 0 ? " · " + myObjects.map((o) => o.name).join(", ") : tr(" · nema dodijeljenih objekata")}
        </div>
        {Object.keys(grandTotals).length > 0 && (
          <div style={{ marginTop: 10, fontSize: 14, fontWeight: 700, color: S.green }}>
            {tr("Ukupno zarađeno (sva vidljiva povijest):")}{" "}{Object.entries(grandTotals).map(([c, v]) => money(v, c)).join(" + ")}
          </div>
        )}
        {Object.keys(paidTotals).length > 0 && (
          <div style={{ marginTop: 4, fontSize: 13, color: S.sub }}>
            {tr("Ukupno mu isplaćeno:")}{" "}{Object.entries(paidTotals).map(([c, v]) => money(v, c)).join(" + ")}
          </div>
        )}
        {Object.values(owedTotals).some((v) => Math.abs(v) > 0.004) && (
          <div style={{ marginTop: 8, padding: "8px 10px", borderRadius: 10, background: Object.values(owedTotals).some((v) => v > 0) ? S.redSoft : S.greenSoft }}>
            {Object.entries(owedTotals).filter(([, v]) => Math.abs(v) > 0.004).map(([c, v]) => (
              <div key={c} style={{ fontWeight: 800, fontSize: 14, color: v > 0 ? S.red : S.green }}>
                {v > 0 ? tr("Dugujem mu: {0}", money(v, c)) : tr("Preplaćeno: {0}", money(-v, c))}
              </div>
            ))}
          </div>
        )}
      </Card>

      {profile.role !== "admin" && (
        <Card style={{ marginTop: 10 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("💸 Upiši isplatu (koliko sam mu dao)")}</div>
          <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 6 }}>
            <input inputMode="decimal" value={payForm.amount} onChange={(e) => setPayForm((f) => ({ ...f, amount: e.target.value }))}
              placeholder={tr("Iznos")} style={{ flex: 1 }} />
            <CurChips small value={payForm.cur} onChange={(v) => setPayForm((f) => ({ ...f, cur: v }))} />
          </div>
          <input type="date" value={payForm.date} onChange={(e) => setPayForm((f) => ({ ...f, date: e.target.value }))} style={{ marginBottom: 6 }} />
          <input value={payForm.note} onChange={(e) => setPayForm((f) => ({ ...f, note: e.target.value }))} placeholder={tr("Napomena (opcionalno)")} style={{ marginBottom: 8 }} />
          <Btn small onClick={savePayout} style={{ width: "100%" }}>{tr("Spremi isplatu")}</Btn>

          {myPayouts.length > 0 && (
            <div style={{ marginTop: 10, borderTop: `1px solid ${S.line}`, paddingTop: 6 }}>
              {myPayouts.map((c) => (
                <div key={c.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", fontSize: 13 }}>
                  <span style={{ flex: 1 }}>
                    {fmtDate(c.payDate)}{c.note ? " · " + c.note : ""}
                  </span>
                  <span className="num" style={{ fontWeight: 700 }}>{money(c.amount, c.currency)}</span>
                  <button onClick={() => api.delCommissionPayout(c, profile.name)} title={tr("Obriši")} style={{ background: "none", border: "none", color: S.red, fontSize: 15, cursor: "pointer", padding: 2 }}>✕</button>
                </div>
              ))}
            </div>
          )}
        </Card>
      )}

      {profile.role === "admin" ? (
        <Card style={{ marginTop: 10 }}>
          <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Admin nema proviziju po satu.")}</div>
        </Card>
      ) : monthRows.length === 0 ? (
        <Card style={{ marginTop: 10 }}>
          <Empty text={tr("Još nema odrađenih sati na dodijeljenim objektima.")} />
        </Card>
      ) : (
        monthRows.map((m) => (
          <Card key={m.mk} style={{ marginTop: 10 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <div style={{ fontWeight: 800 }}>{monthLabel(m.mk)}</div>
              <div className="num" style={{ fontWeight: 700, color: S.green }}>
                {Object.entries(m.totalsByCur).map(([c, v]) => money(v, c)).join(" + ")}
              </div>
            </div>
            {m.perObj.map((r) => (
              <div key={r.obj.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "4px 0", fontSize: 13, borderTop: `1px solid ${S.line}` }}>
                <span style={{ flex: 1 }}>{r.obj.name}</span>
                <span className="num" style={{ color: S.sub, marginRight: 8 }}>{fmtH(r.hrs)} × {money(r.rate, r.cur)}</span>
                <span className="num" style={{ fontWeight: 700 }}>{money(r.amount, r.cur)}</span>
              </div>
            ))}
          </Card>
        ))
      )}
    </>
  );
}

/* ================================================================== */
/*  DETALJ RADNIKA                                                     */
/* ================================================================== */
function WorkerDetail({ worker, data, api, onBack }) {
  const [edit, setEdit] = useState(false);
  const [form, setForm] = useState({ ...worker, rate: String(worker.rate || ""), rateCur: wCur(worker), ...splitNameForEdit(worker.name) });
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
      if (f.size > 10 * 1024 * 1024) { setDocErr(tr("\"{0}\" je veća od 10 MB.", f.name)); continue; }
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
    if (error) { setDocErr(tr("Brisanje dokumenata može samo admin.")); return; }
    await api.logDoc(`Obrisao dokument radnika ${worker.name}`);
    setViewer(null);
    await loadDocs();
  };

  const save = async () => {
    const fullName = `${form.lastName.trim()} ${form.firstName.trim()}`.trim();
    if (await api.updWorker(worker.id, { ...form, name: fullName }, fullName)) setEdit(false);
  };
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
      <button onClick={onBack} style={{ background: "none", border: "none", color: S.green, fontWeight: 700, fontSize: 14.5, padding: "2px 0 12px", cursor: "pointer" }}>{tr("← Natrag")}</button>

      <Card>
        {!edit ? (
          <>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ fontSize: 19, fontWeight: 800 }}>{worker.name}</div>
                <div style={{ color: S.sub, fontSize: 13.5, marginTop: 3 }}>{worker.archived ? tr("📁 Bivši radnik · ") : ""}{objName(worker.objectId) || tr("Bez objekta")}{worker.position ? " · " + worker.position : ""}</div>
              </div>
              {rateNow(data, worker) > 0
                ? <Tag color={S.green} bg={S.greenSoft}>{money(rateNow(data, worker), wCur(worker))}/h</Tag>
                : <Tag color={S.amber} bg={S.amberSoft}>{tr("bez satnice")}</Tag>}
            </div>
            <div style={{ marginTop: 10, fontSize: 14.5 }}>📞 {worker.phone ? <a href={"tel:" + worker.phone} style={{ color: S.ink }}>{worker.phone}</a> : "—"}</div>
            {(worker.permitExpiry || worker.contractExpiry) && (
              <div className="num" style={{ marginTop: 6, fontSize: 13, color: S.sub }}>
                {worker.permitExpiry && <>{tr("Radna dozvola do")}{" "}{fmtDate(worker.permitExpiry)}</>}
                {worker.permitExpiry && worker.contractExpiry && " · "}
                {worker.contractExpiry && <>{tr("Ugovor do")}{" "}{fmtDate(worker.contractExpiry)}</>}
              </div>
            )}
            {worker.note && <div style={{ marginTop: 8, fontSize: 13.5, color: S.sub, whiteSpace: "pre-wrap" }}>{worker.note}</div>}
            <div style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap" }}>
              <Btn small kind="ghost" onClick={() => { setForm({ ...worker, rate: String(worker.rate || ""), rateCur: wCur(worker) }); setEdit(true); }}>{tr("Uredi")}</Btn>
              <Btn small kind="ghost" onClick={() => api.archiveWorker(worker, !worker.archived)} style={{ color: worker.archived ? S.green : S.amber, fontWeight: 700 }}>
                {worker.archived ? tr("↩ Vrati u aktivne") : tr("📁 Završio s radom")}
              </Btn>
              {api.admin ? (
                <>
                  <Btn small kind={confirmRemove ? "primary" : "danger"} onClick={remove} style={confirmRemove ? { background: S.red } : undefined}>
                    {confirmRemove ? tr("Potvrdi brisanje") : tr("Obriši")}
                  </Btn>
                  {confirmRemove && <Btn small kind="ghost" onClick={() => setConfirmRemove(false)}>{tr("Odustani")}</Btn>}
                </>
              ) : <span style={{ fontSize: 12, color: S.sub, alignSelf: "center" }}>{tr("🔒 brisanje može samo admin")}</span>}
            </div>
          </>
        ) : (
          <>
            <div style={{ display: "flex", gap: 10 }}>
              <div style={{ flex: 1 }}>
                <Field label={tr("Prezime")}><input value={form.lastName} onChange={(e) => setForm({ ...form, lastName: e.target.value })} /></Field>
              </div>
              <div style={{ flex: 1 }}>
                <Field label={tr("Ime")}><input value={form.firstName} onChange={(e) => setForm({ ...form, firstName: e.target.value })} /></Field>
              </div>
            </div>
            <Field label={tr("Broj telefona")}><input value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
            <Field label={tr("Satnica ({0} / sat)", form.rateCur === "CZK" ? "Kč" : "€")}>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input value={form.rate} inputMode="decimal" onChange={(e) => setForm({ ...form, rate: e.target.value })} />
                <CurChips small value={form.rateCur} onChange={(v) => setForm({ ...form, rateCur: v })} />
              </div>
            </Field>
            {isRateSuspicious(form.rate, form.rateCur) && (
              <div style={{ fontSize: 12.5, color: S.amber, fontWeight: 600, margin: "-6px 0 10px" }}>
                {tr("⚠️ Ova satnica izgleda neobično — provjeri nisi li se utipkala.")}
              </div>
            )}
            <Field label={tr("Glavni objekt")}><ObjectSelect data={data} api={api} value={form.objectId || ""} onChange={(v) => setForm({ ...form, objectId: v })} /></Field>
            <Field label={tr("Pozicija (npr. HSK, kuhinja, bar…)")}><input value={form.position || ""} onChange={(e) => setForm({ ...form, position: e.target.value })} placeholder={tr("npr. HSK")} /></Field>
            <div style={{ display: "flex", gap: 10 }}>
              <div style={{ flex: 1 }}><Field label={tr("Istek radne dozvole")}><input type="date" value={form.permitExpiry || ""} onChange={(e) => setForm({ ...form, permitExpiry: e.target.value })} /></Field></div>
              <div style={{ flex: 1 }}><Field label={tr("Istek ugovora")}><input type="date" value={form.contractExpiry || ""} onChange={(e) => setForm({ ...form, contractExpiry: e.target.value })} /></Field></div>
            </div>
            {api.admin && <Field label={tr("Broj bankovnog računa (Číslo účtu) — za knjigovođu")}><input value={form.bankAccount || ""} onChange={(e) => setForm({ ...form, bankAccount: e.target.value })} placeholder={tr("npr. 123456789/0800")} /></Field>}
            <Field label={tr("Napomena")}><textarea rows={2} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
            <div style={{ display: "flex", gap: 8 }}>
              <Btn small onClick={save}>{tr("Spremi")}</Btn>
              <Btn small kind="ghost" onClick={() => setEdit(false)}>{tr("Odustani")}</Btn>
            </div>
          </>
        )}
      </Card>

      <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA" }}>
        <div style={{ fontWeight: 700, color: S.blue, marginBottom: 8 }}>{tr("🏨 Objekt:")}{" "}{objName(worker.objectId) || tr("Bez objekta")}</div>
        <div style={{ display: "flex", gap: 8, alignItems: "flex-end", flexWrap: "wrap" }}>
          <div style={{ flex: "1 1 150px" }}><Field label={tr("Prebaci na")}><ObjectSelect data={data} api={api} value={moveObj} onChange={setMoveObj} /></Field></div>
          <div style={{ flex: "1 1 120px" }}><Field label={tr("Od datuma")}><input type="date" value={moveDate} onChange={(e) => setMoveDate(e.target.value)} /></Field></div>
          <div style={{ marginBottom: 10 }}><Btn small onClick={transfer}>{tr("Prebaci")}</Btn></div>
        </div>
        {myAssignments.length > 0 && (
          <div>
            <div style={{ fontSize: 12.5, fontWeight: 700, color: S.blue, marginBottom: 4 }}>{tr("Povijest premještaja")}</div>
            {myAssignments.map((a, i) => (
              <div key={a.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", borderBottom: `1px solid #CBDCEA`, fontSize: 13.5 }}>
                <span className="num" style={{ flex: 1 }}>
                  {tr("od")}{" "}{fmtDate(a.from)} · <b>{objName(a.objectId) || tr("Bez objekta")}</b>
                  {i === 0 ? <span style={{ color: S.green, fontWeight: 700 }}> {tr("(trenutno)")}</span> : ""}
                </span>
                <button onClick={() => api.delAssignment(a, worker.name)} style={{ background: "none", border: "none", color: S.red, fontSize: 15, cursor: "pointer", padding: 4 }}>✕</button>
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card style={{ background: S.greenSoft, borderColor: "#C5DED2" }}>
        <div style={{ fontWeight: 700, color: S.green, marginBottom: 8 }}>
          {tr("💶 Satnica:")}{" "}{rateNow(data, worker) > 0 ? money(rateNow(data, worker), wCur(worker)) + "/h" : tr("nije unesena — dodaj je ovdje")}
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "flex-end", flexWrap: "wrap" }}>
          <div style={{ flex: "1 1 120px" }}><Field label={tr("Nova satnica ({0}/h)", wCur(worker) === "CZK" ? "Kč" : "€")}><input inputMode="decimal" value={rateNew} onChange={(e) => setRateNew(e.target.value)} onKeyDown={(e) => e.key === "Enter" && changeRate()} placeholder={wCur(worker) === "CZK" ? tr("npr. 160") : tr("npr. 7.00")} /></Field></div>
          <div style={{ flex: "1 1 130px" }}><Field label={tr("Vrijedi od")}><input type="date" value={rateFrom} onChange={(e) => setRateFrom(e.target.value)} /></Field></div>
          <div style={{ marginBottom: 10 }}><Btn small onClick={changeRate}>{tr("Promijeni")}</Btn></div>
        </div>
        {isRateSuspicious(rateNew, wCur(worker)) && (
          <div style={{ fontSize: 12.5, color: S.amber, fontWeight: 600, margin: "-6px 0 10px" }}>
            {tr("⚠️ Ova satnica izgleda neobično — provjeri nisi li se utipkala.")}
          </div>
        )}
        <div style={{ fontSize: 12.5, fontWeight: 700, color: S.green, marginBottom: 2 }}>{tr("Povijest satnice")}</div>
        <div className="num" style={{ fontSize: 13.5, padding: "4px 0", color: S.sub }}>{tr("početna ·")}{" "}<b>{worker.rate > 0 ? money(worker.rate, wCur(worker)) + "/h" : "—"}</b></div>
        {myRates.map((r) => (
          <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "4px 0", borderTop: `1px solid #C5DED2`, fontSize: 13.5 }}>
            <span className="num" style={{ flex: 1 }}>{tr("od")}{" "}{fmtDate(r.from)} · <b>{money(r.rate, wCur(worker))}/h</b></span>
            <button onClick={() => api.delRateChange(r, worker.name)} style={{ background: "none", border: "none", color: S.red, fontSize: 15, cursor: "pointer", padding: 4 }}>✕</button>
          </div>
        ))}
      </Card>

      {(() => {
        const y = new Date().getFullYear();
        const ytdRows = calcRows(data, (d) => (d || "").slice(0, 4) === String(y), new Set());
        const r = ytdRows.find((x) => x.w.id === worker.id);
        if (!r || (r.hours === 0 && r.pays.length === 0)) return null;
        return (
          <Card style={{ background: S.greenSoft, borderColor: "#C5DED2" }}>
            <div style={{ fontWeight: 700, color: S.green, marginBottom: 6 }}>{tr("📆 Ukupno u")}{" "}{y}.</div>
            <div className="num" style={{ fontSize: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Sati")}</span><b>{fmtH(r.hours)}</b></div>
              {(r.net !== 0 || r.czk.net === 0) && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Zarada (€)")}</span><b>{eur(r.net)}</b></div>}
              {r.czk.net !== 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Zarada (Kč)")}</span><b>{czk(r.czk.net)}</b></div>}
            </div>
          </Card>
        );
      })()}

      <Card>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <div style={{ fontWeight: 700 }}>{tr("Dokumenti i ugovori")}</div>
          <Btn small kind="ghost" onClick={() => fileRef.current?.click()}>{tr("+ Slika ili datoteka")}</Btn>
          <input ref={fileRef} type="file" accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.txt" multiple onChange={uploadDoc} style={{ display: "none" }} />
        </div>
        {docErr && <div style={{ color: S.red, fontSize: 13, marginBottom: 8 }}>{docErr}</div>}
        {docs === null ? <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Učitavam…")}</div>
          : docs.length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Slikaj ugovor ili dodaj PDF.")}</div>
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
        <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("Zadnji upisi sati")}</div>
        {logs.length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Još nema upisanih sati.")}</div>
          : logs.map((l) => (
            <div key={l.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderBottom: `1px solid ${S.line}`, fontSize: 14 }}>
              <span style={{ flex: 1 }}>{fmtDate(l.date)} · <LogSpanLabel l={l} />{objName(l.objectId) ? " · " + objName(l.objectId) : ""}</span>
              <span className="num" style={{ fontWeight: 700, color: !l.monthly && isDayHoursSuspicious(l.hours) ? S.red : "inherit" }}>
                {!l.monthly && isDayHoursSuspicious(l.hours) && "⚠️ "}{fmtH(l.hours)}
              </span>
              {(api.admin || l.createdBy === api.uid()) && (
                <button onClick={() => api.delLog(l, worker.name)} title={tr("Obriši")} style={{ background: "none", border: "none", color: S.red, fontSize: 16, cursor: "pointer", padding: 4 }}>✕</button>
              )}
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
            <a href={viewer.url} target="_blank" rel="noreferrer" style={{ background: S.green, color: "#fff", borderRadius: 10, padding: "9px 14px", fontWeight: 700, textDecoration: "none", fontSize: 13.5 }}>{tr("⬇ Otvori / preuzmi")}</a>
            {api.admin && <Btn small kind="ghost" onClick={(e) => { e.stopPropagation(); delDoc({ name: viewer.name }); }}>{tr("Obriši")}</Btn>}
            <Btn small onClick={() => setViewer(null)}>{tr("Zatvori")}</Btn>
          </div>
        </div>
      )}
    </>
  );
}

/* ================================================================== */
/*  OBJEKT — BRZI UNOS + kopiranje dana                                */
/* ================================================================== */
/* ================================================================== */
/*  BRIGADA (jednodnevni ljudi plaćeni u gotovini) + GOTOVINA KOD      */
/*  KOORDINATORA. Tablice: brigade_costs (paid_by = koordinator koji  */
/*  je platio iz gotovine, null = admin), cash_advances (dano/vraćeno) */
/* ================================================================== */
// saldo gotovine po koordinatoru i valuti: dano − vraćeno − isplaćeno brigadama
// plus = koordinator još drži tvoj novac, minus = ti njemu duguješ
function cashBalances(data) {
  const m = new Map();
  const get = (pid, cur) => {
    if (!m.has(pid)) m.set(pid, { EUR: { given: 0, returned: 0, paid: 0, bal: 0 }, CZK: { given: 0, returned: 0, paid: 0, bal: 0 } });
    return m.get(pid)[cur === "EUR" ? "EUR" : "CZK"];
  };
  (data.cashAdvances || []).forEach((c) => { const r = get(c.profileId, c.currency); r[c.kind] = round2(r[c.kind] + c.amount); });
  (data.brigadeCosts || []).filter((b) => b.paidBy).forEach((b) => { const r = get(b.paidBy, b.currency); r.paid = round2(r.paid + b.amount); });
  m.forEach((v) => ["EUR", "CZK"].forEach((c) => { v[c].bal = round2(v[c].given - v[c].returned - v[c].paid); }));
  return m;
}
const cashBalFor = (data, pid) => cashBalances(data).get(pid) || null;
// koliko hotel plati za tu brigadu (posebna cijena ili zadana cijena objekta)
const brigadeRevenue = (object, b) => {
  const rate = b.billRate > 0 ? b.billRate : (object?.billRate || 0);
  return { amount: round2(b.hours * rate), cur: b.billRate > 0 ? b.billCur : (object?.billCur || "EUR") };
};
// zarada brigade u valuti troška (preračun po tečaju iz postavki ako hotel plaća u drugoj valuti)
const brigadeProfit = (data, object, b) => {
  const rv = brigadeRevenue(object, b);
  const kurs = Number(data.settings?.czk_rate) || 25;
  const rev = rv.cur === b.currency ? rv.amount : b.currency === "CZK" ? rv.amount * kurs : rv.amount / kurs;
  return round2(rev - b.amount);
};

function BrigadeCard({ object, data, api, monthBrigade, defaultDate }) {
  const coordinators = useMemo(() => {
    const onObj = new Set((data.objectMembers || []).filter((m) => m.object_id === object.id).map((m) => m.user_id));
    return [...(data.profiles || [])].filter((p) => p.role !== "admin")
      .sort((a, b) => (onObj.has(b.id) - onObj.has(a.id)) || (a.name || "").localeCompare(b.name || "", "hr"));
  }, [data.profiles, data.objectMembers, object.id]);
  // zadana cash satnica = satnica zadnje brigade na ovom objektu (koliko si dao po satu)
  const last = useMemo(() => [...(data.brigadeCosts || [])].filter((b) => b.objectId === object.id && b.hours > 0)
    .sort((a, b) => b.date.localeCompare(a.date))[0], [data.brigadeCosts, object.id]);
  const blank = () => ({
    date: defaultDate || todayISO(), people: "1", hoursPer: "", payRate: last ? String(round2(last.amount / last.hours)) : "",
    amount: "", amountTouched: false, cur: last ? last.currency : (object.billCur === "EUR" ? "EUR" : "CZK"),
    names: "", note: "", billRate: "", billCur: object.billCur || "EUR", paidBy: "", more: false,
  });
  const [f, setF] = useState(blank);
  useEffect(() => { setF((p) => ({ ...p, date: defaultDate || p.date })); }, [defaultDate]);
  const people = Math.max(1, Math.round(parseNum(f.people) || 1));
  const hours = round2(people * (parseNum(f.hoursPer) || 0));
  const autoAmount = round2(hours * (parseNum(f.payRate) || 0));
  const amount = f.amountTouched ? round2(parseNum(f.amount) || 0) : autoAmount;
  // upisan ukupni iznos → satnica se računa iz njega (iznos / sati)
  const effRate = f.amountTouched ? (hours > 0 ? round2(amount / hours) : 0) : round2(parseNum(f.payRate) || 0);
  const preview = { hours, amount, currency: f.cur, billRate: round2(parseNum(f.billRate) || 0), billCur: f.billCur };
  const rv = brigadeRevenue(object, preview);
  const profit = brigadeProfit(data, object, preview);
  const payer = coordinators.find((p) => p.id === f.paidBy);
  const payerBal = payer ? (cashBalFor(data, payer.id) || { EUR: { bal: 0 }, CZK: { bal: 0 } })[f.cur].bal : 0;

  const save = async () => {
    if (!f.date || hours <= 0 || amount <= 0) return;
    const ok = await api.addBrigadeCost(object, { date: f.date, hours, amount, cur: f.cur, note: f.note.trim(), names: f.names.trim(),
      people, billRate: preview.billRate, billCur: f.billCur, paidBy: f.paidBy || null }, payer?.name || "");
    if (ok) setF((p) => ({ ...blank(), date: p.date, payRate: effRate ? String(effRate) : p.payRate, cur: p.cur, paidBy: p.paidBy, people: p.people, hoursPer: p.hoursPer }));
  };
  // "Ponovi": ista brigada (ljudi, sati, iznos, tko plaća) — samo promijeniš datum i spremiš
  const repeat = (bg) => setF({
    ...blank(), date: defaultDate || todayISO(), people: String(bg.people || 1), hoursPer: String(round2(bg.hours / (bg.people || 1))),
    payRate: String(round2(bg.amount / bg.hours)), amount: String(bg.amount), amountTouched: false, cur: bg.currency,
    names: bg.names, note: bg.note, billRate: bg.billRate > 0 ? String(bg.billRate) : "", billCur: bg.billCur, paidBy: bg.paidBy || "", more: bg.billRate > 0,
  });
  const pName = (id) => (data.profiles || []).find((p) => p.id === id)?.name || tr("(obrisan)");
  const list = [...monthBrigade].sort((a, b) => b.date.localeCompare(a.date));
  const totE = round2(list.filter((b) => b.currency === "EUR").reduce((s, b) => s + b.amount, 0));
  const totK = round2(list.filter((b) => b.currency === "CZK").reduce((s, b) => s + b.amount, 0));
  const totH = round2(list.reduce((s, b) => s + b.hours, 0));

  return (
    <Card>
      <div style={{ fontWeight: 800, fontSize: 15, marginBottom: 4 }}>{tr("👷 Brigada (bez imena)")}</div>
      <div style={{ fontSize: 11.5, color: S.sub, marginBottom: 10 }}>
        {tr("Ljudi koje pošalješ na dan i odmah platiš u gotovini. Sati se naplaćuju hotelu, a iznos ide kao trošak objekta.")}
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 6 }}>
        <div style={{ flex: "1 1 130px" }}><Field label={tr("Datum")}>
          <input type="date" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} />
        </Field></div>
        <div style={{ flex: "0 1 80px" }}><Field label={tr("Ljudi")}>
          <input inputMode="numeric" value={f.people} onChange={(e) => setF({ ...f, people: e.target.value })} placeholder="1" />
        </Field></div>
        <div style={{ flex: "0 1 90px" }}><Field label={tr("Sati / osobi")}>
          <input inputMode="decimal" value={f.hoursPer} onChange={(e) => setF({ ...f, hoursPer: e.target.value })} placeholder="8" />
        </Field></div>
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 6, alignItems: "flex-end" }}>
        <div style={{ flex: "1 1 110px" }}><Field label={tr("Daješ po satu")}>
          <input inputMode="decimal" value={f.amountTouched ? (effRate ? String(effRate) : "") : f.payRate}
            onChange={(e) => setF({ ...f, payRate: e.target.value, amountTouched: false })} placeholder="0" />
        </Field></div>
        <div style={{ flex: "1 1 110px" }}><Field label={tr("Ukupno isplaćeno")}>
          <input inputMode="decimal" value={f.amountTouched ? f.amount : (autoAmount ? String(autoAmount) : "")}
            onChange={(e) => setF({ ...f, amount: e.target.value, amountTouched: true })} placeholder={tr("iznos")} />
        </Field></div>
        <div style={{ marginBottom: 10 }}><CurChips small value={f.cur} onChange={(v) => setF({ ...f, cur: v })} /></div>
      </div>
      <Field label={tr("Tko je platio")}>
        <select value={f.paidBy} onChange={(e) => setF({ ...f, paidBy: e.target.value })}>
          <option value="">{tr("Ja (iz svog džepa)")}</option>
          {coordinators.map((p) => <option key={p.id} value={p.id}>{tr("Koordinator: {0}", p.name || tr("(bez imena)"))}</option>)}
        </select>
      </Field>
      {payer && (
        <div className="num" style={{ fontSize: 12, margin: "-4px 0 10px", fontWeight: 600, color: payerBal - amount < 0 ? S.red : S.sub }}>
          {tr("Kod {0} sad: {1} → nakon ove brigade: {2}", payer.name, money(payerBal, f.cur), money(round2(payerBal - amount), f.cur))}
          {payerBal - amount < 0 && <> · {tr("ti ćeš mu dugovati")}</>}
        </div>
      )}
      <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
        <input value={f.names} onChange={(e) => setF({ ...f, names: e.target.value })} placeholder={tr("Imena (opcionalno, npr. Ivan, Petar)")} style={{ flex: 1 }} />
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
        <input value={f.note} onChange={(e) => setF({ ...f, note: e.target.value })} placeholder={tr("Napomena (opcionalno)")} style={{ flex: 1 }} />
      </div>
      {f.more ? (
        <>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 6, alignItems: "center" }}>
            <div style={{ flex: "1 1 160px" }}>
              <input inputMode="decimal" value={f.billRate} onChange={(e) => setF({ ...f, billRate: e.target.value })}
                placeholder={tr("cijena/sat (zadano {0})", object.billRate || 0)} />
            </div>
            <CurChips small value={f.billCur} onChange={(v) => setF({ ...f, billCur: v })} />
          </div>
          <div style={{ fontSize: 11, color: S.sub, margin: "-2px 0 8px" }}>
            {tr("Ostavi prazno da koristi zadanu cijenu objekta (")}{money(object.billRate || 0, object.billCur || "EUR")}{tr("/h) — upiši samo ako hotel ovu brigadu plaća drugačije.")}
          </div>
        </>
      ) : (
        <button onClick={() => setF({ ...f, more: true })} style={{ background: "none", border: "none", color: S.blue, fontSize: 12, cursor: "pointer", padding: "0 0 8px" }}>
          {tr("+ posebna cijena za hotel")}
        </button>
      )}
      {hours > 0 && (
        <div className="num" style={{ fontSize: 12.5, background: S.greenSoft, borderRadius: 8, padding: "7px 10px", marginBottom: 8 }}>
          {people > 1 ? `${people} × ${fmtH(round2(hours / people))} = ` : ""}<b>{fmtH(hours)}</b>
          {" · "}{tr("hotel plaća")} <b>{money(rv.amount, rv.cur)}</b>
          {amount > 0 && <> · {tr("isplaćeno")} <b>{money(amount, f.cur)}</b>{effRate ? ` (${money(effRate, f.cur)}/h)` : ""} · {tr("zarada")} <b style={{ color: profit < 0 ? S.red : S.green }}>{money(profit, f.cur)}</b></>}
        </div>
      )}
      <Btn small onClick={save} disabled={hours <= 0 || amount <= 0} style={{ width: "100%", marginBottom: 10 }}>{tr("Spremi brigadu")}</Btn>

      {list.length > 0 && (
        <div>
          <div className="num" style={{ fontSize: 12, color: S.sub, fontWeight: 700, marginBottom: 4 }}>
            {tr("Ovaj mjesec: {0} upisa · {1}", list.length, fmtH(totH))}{totE ? " · " + eur(totE) : ""}{totK ? " · " + czk(totK) : ""}
          </div>
          {list.map((bg) => {
            const pr = brigadeProfit(data, object, bg);
            return (
              <div key={bg.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", fontSize: 12.5, borderBottom: "1px dashed #EBD9B4" }}>
                <span style={{ flex: 1, minWidth: 0 }}>
                  <b>{fmtDate(bg.date)}</b> · {bg.people > 1 ? `${bg.people} × ${fmtH(round2(bg.hours / bg.people))}` : fmtH(bg.hours)}
                  {bg.names ? " · " + bg.names : ""}{bg.note ? " · " + bg.note : ""}
                  {bg.billRate > 0 && <span style={{ color: S.blue, fontWeight: 700 }}> · {money(bg.billRate, bg.billCur)}{tr("/h posebno")}</span>}
                  <div style={{ fontSize: 11, color: S.sub }}>
                    {bg.paidBy ? tr("platio {0} iz gotovine", pName(bg.paidBy)) : tr("platio ja")}
                    {bg.hours > 0 ? ` · ${money(round2(bg.amount / bg.hours), bg.currency)}/h` : ""}
                    {" · "}{tr("zarada")} <span className="num" style={{ color: pr < 0 ? S.red : S.green, fontWeight: 700 }}>{money(pr, bg.currency)}</span>
                  </div>
                </span>
                <span className="num" style={{ fontWeight: 700 }}>{money(bg.amount, bg.currency)}</span>
                <button title={tr("Ponovi")} onClick={() => repeat(bg)} style={{ background: "none", border: "none", color: S.blue, fontSize: 14, cursor: "pointer", padding: 2 }}>↻</button>
                <button onClick={() => window.confirm(tr("Obrisati ovu brigadu?")) && api.delBrigadeCost(bg, object.name)} style={{ background: "none", border: "none", color: S.red, fontSize: 14, cursor: "pointer", padding: 2 }}>✕</button>
              </div>
            );
          })}
        </div>
      )}
    </Card>
  );
}

function CashPanel({ data, api }) {
  const coordinators = [...(data.profiles || [])].filter((p) => p.role !== "admin").sort((a, b) => (a.name || "").localeCompare(b.name || "", "hr"));
  const bal = cashBalances(data);
  const [form, setForm] = useState({ profileId: "", kind: "given", amount: "", cur: "CZK", date: todayISO(), note: "" });
  const [open, setOpen] = useState(null);
  const oName = (id) => data.objects.find((o) => o.id === id)?.name || "";
  const save = async () => {
    const p = coordinators.find((x) => x.id === form.profileId);
    const amt = round2(parseNum(form.amount) || 0);
    if (!p || amt <= 0) return;
    if (await api.addCashAdvance(p, form.kind, amt, form.cur, form.date, form.note.trim())) setForm({ ...form, amount: "", note: "" });
  };
  const active = coordinators.filter((p) => bal.has(p.id));
  const balLine = (b, c) => {
    const v = b[c].bal;
    if (!b[c].given && !b[c].returned && !b[c].paid) return null;
    return (
      <div key={c} className="num" style={{ fontSize: 13, fontWeight: 700, color: v > 0 ? S.amber : v < 0 ? S.red : S.green }}>
        {v > 0 ? tr("Kod njega: {0}", money(v, c)) : v < 0 ? tr("Ti mu duguješ: {0}", money(-v, c)) : tr("Na nuli ({0})", c === "CZK" ? "Kč" : "€")}
      </div>
    );
  };

  return (
    <div style={{ marginTop: 8 }}>
      <Card>
        <div style={{ fontWeight: 700, marginBottom: 4 }}>{tr("💵 Dao sam / vratio mi gotovinu")}</div>
        <div style={{ fontSize: 11.5, color: S.sub, marginBottom: 10 }}>
          {tr("Kad koordinatoru na hotelu daš novac za brigade, upiši ga ovdje. Brigade koje on plati (\"Tko je platio\" kod brigade) skidaju se s tog iznosa.")}
        </div>
        {coordinators.length === 0 ? <div style={{ fontSize: 12.5, color: S.sub }}>{tr("Još nema zaposlenika.")}</div> : (<>
          <Field label={tr("Koordinator")}>
            <select value={form.profileId} onChange={(e) => setForm({ ...form, profileId: e.target.value })}>
              <option value="">{tr("— odaberi —")}</option>
              {coordinators.map((p) => <option key={p.id} value={p.id}>{p.name || tr("(bez imena)")}</option>)}
            </select>
          </Field>
          <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
            {[["given", tr("Dao sam mu")], ["returned", tr("Vratio mi je")]].map(([k, l]) => (
              <button key={k} onClick={() => setForm({ ...form, kind: k })} style={{ flex: 1, padding: "7px 8px", borderRadius: 8, fontWeight: 700, fontSize: 13, cursor: "pointer",
                background: form.kind === k ? S.green : "#fff", color: form.kind === k ? "#fff" : S.sub, border: `1px solid ${form.kind === k ? S.green : S.line}` }}>{l}</button>
            ))}
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginBottom: 8 }}>
            <input inputMode="decimal" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} placeholder={tr("iznos")} style={{ flex: "1 1 110px" }} />
            <CurChips small value={form.cur} onChange={(v) => setForm({ ...form, cur: v })} />
            <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} style={{ flex: "1 1 130px" }} />
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <input value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} placeholder={tr("Napomena (opcionalno)")} style={{ flex: 1 }} />
            <Btn small onClick={save} disabled={!form.profileId || !(parseNum(form.amount) > 0)}>{tr("Spremi")}</Btn>
          </div>
        </>)}
      </Card>

      {active.length === 0 ? <Empty text={tr("Još nema gotovine kod koordinatora.")} /> : active.map((p) => {
        const b = bal.get(p.id);
        const isOpen = open === p.id;
        const events = [
          ...(data.cashAdvances || []).filter((c) => c.profileId === p.id).map((c) => ({ id: "c" + c.id, date: c.date, cur: c.currency,
            delta: c.kind === "given" ? c.amount : -c.amount, label: c.kind === "given" ? tr("Dao sam mu") : tr("Vratio mi je"), note: c.note, rec: c })),
          ...(data.brigadeCosts || []).filter((bg) => bg.paidBy === p.id).map((bg) => ({ id: "b" + bg.id, date: bg.date, cur: bg.currency, delta: -bg.amount,
            label: tr("Brigada {0}", oName(bg.objectId)) + " · " + (bg.people > 1 ? `${bg.people} × ${fmtH(round2(bg.hours / bg.people))}` : fmtH(bg.hours)), note: bg.names || bg.note })),
        ].sort((x, y) => x.date.localeCompare(y.date) || (y.delta - x.delta));
        const run = { EUR: 0, CZK: 0 };
        events.forEach((e) => { run[e.cur] = round2(run[e.cur] + e.delta); e.after = run[e.cur]; });
        return (
          <Card key={p.id} style={{ cursor: "pointer" }} onClick={() => setOpen(isOpen ? null : p.id)}>
            <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 800 }}>{p.name || tr("(bez imena)")}</div>
                {["CZK", "EUR"].map((c) => (b[c].given || b[c].returned || b[c].paid) ? (
                  <div key={c} className="num" style={{ fontSize: 11.5, color: S.sub }}>
                    {tr("dano {0} · brigade {1} · vraćeno {2}", money(b[c].given, c), money(b[c].paid, c), money(b[c].returned, c))}
                  </div>
                ) : null)}
              </div>
              <div style={{ textAlign: "right" }}>{["CZK", "EUR"].map((c) => balLine(b, c))}</div>
            </div>
            {isOpen && (
              <div style={{ marginTop: 10 }} onClick={(e) => e.stopPropagation()}>
                {[...events].reverse().map((e) => (
                  <div key={e.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", fontSize: 12.5, borderTop: `1px dashed ${S.line}` }}>
                    <span style={{ flex: 1, minWidth: 0 }}>
                      <b>{fmtDate(e.date)}</b> · {e.label}{e.note ? " · " + e.note : ""}
                      <div className="num" style={{ fontSize: 11, color: S.sub }}>{tr("stanje nakon: {0}", money(e.after, e.cur))}</div>
                    </span>
                    <span className="num" style={{ fontWeight: 700, color: e.delta >= 0 ? S.green : S.red }}>{e.delta >= 0 ? "+" : "−"}{money(Math.abs(e.delta), e.cur)}</span>
                    {e.rec ? (
                      <button onClick={() => window.confirm(tr("Obrisati ovaj zapis?")) && api.delCashAdvance(e.rec, p.name)} style={{ background: "none", border: "none", color: S.red, fontSize: 14, cursor: "pointer", padding: 2 }}>✕</button>
                    ) : <span style={{ width: 18 }} />}
                  </div>
                ))}
                <div style={{ fontSize: 11, color: S.sub, marginTop: 6 }}>{tr("Brigade se brišu i mijenjaju na objektu (kartica Sati).")}</div>
              </div>
            )}
          </Card>
        );
      })}
    </div>
  );
}

function ObjectDetail({ object, data, api, onBack, onOpenWorker }) {
  const [mode, setMode] = useState("day");
  const [section, setSection] = useState(api.admin ? "pregled" : "sati"); // pregled | sati | fakture | postavke
  const [volnoMonth, setVolnoMonth] = useState(() => prevMonthKey(curMonth())); // Neplacené volno: zadano prošli mjesec
  const [date, setDate] = useState(todayISO());
  const [month, setMonth] = useState(curMonth());
  const [inputs, setInputs] = useState({});
  const [posEdit, setPosEdit] = useState({});
  const [monthVals, setMonthVals] = useState({}); // workerId -> string, sati za odabrani mjesec
  const [monthSel, setMonthSel] = useState(() => new Set()); // odabrani radnici za grupnu izmjenu
  const [monthBulkVal, setMonthBulkVal] = useState("");
  const [monthSaving, setMonthSaving] = useState(false);
  const [hourEdit, setHourEdit] = useState(null); // { workerId, value } — brzo uređivanje broja sati u sažetku
  const [hourSaving, setHourSaving] = useState(false);
  const [q, setQ] = useState("");
  const [showTransfer, setShowTransfer] = useState(false);
  const [transferSel, setTransferSel] = useState(() => new Set());
  const [transferDate, setTransferDate] = useState(todayISO());
  const [transferBusy, setTransferBusy] = useState(false);
  const [editLog, setEditLog] = useState(null);
  const [rateEdit, setRateEdit] = useState(String(object.billRate || ""));
  const [billCur, setBillCur] = useState(object.billCur || "EUR");
  const [budgetEdit, setBudgetEdit] = useState(String(object.monthlyBudget || ""));
  const [posRateForm, setPosRateForm] = useState({ position: "", mode: "fixed", value: "", currency: object.billCur || "EUR" });
  const [posRateEditing, setPosRateEditing] = useState(null); // pozicija koja se trenutno uređuje, ili null za novu
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
  const [genVat, setGenVat] = useState("");
  const [genVatCur, setGenVatCur] = useState(object.billCur || "EUR");
  const [invFileBusy, setInvFileBusy] = useState(null); // id fakture za koju se trenutno uploada PDF
  const invPdfRef = useRef(null);
  const invPdfTarget = useRef(null);

  const [custNameEdit, setCustNameEdit] = useState(object.invoiceCustomerName || "");
  useEffect(() => { setRateEdit(String(object.billRate || "")); setBillCur(object.billCur || "EUR"); setBudgetEdit(String(object.monthlyBudget || "")); setCustNameEdit(object.invoiceCustomerName || ""); }, [object.id, object.billRate, object.billCur, object.monthlyBudget, object.invoiceCustomerName]);

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
      if (f.size > 15 * 1024 * 1024) { setInvErr(tr("\"{0}\" je veća od 15 MB.", f.name)); continue; }
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

  const objHoursFor = (mo) => round2(
    data.logs.filter((l) => l.objectId === object.id && monthKey(l.date) === mo).reduce((s, l) => s + l.hours, 0)
    + (data.brigadeCosts || []).filter((bg) => bg.objectId === object.id && monthKey(bg.date) === mo).reduce((s, bg) => s + bg.hours, 0)
  );
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
    setGenVat("");
    setGenVatCur(object.billCur || "EUR");
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
    const vatAmount = round2(parseNum(genVat) || 0);
    if (!genNumber.trim() || amount <= 0) return;
    const inv = { objectId: object.id, number: genNumber.trim(), period: genMonth, issueDate: todayISO(), hours, rate, amount, currency: genCur, note: genNote, vatAmount, vatCurrency: genVatCur };
    if (await api.addObjectInvoice(inv, object.name)) {
      printGeneratedInvoice(inv, object, data.settings || {});
      setShowGen(false);
    }
  };
  const myInvoices = (data.objectInvoices || []).filter((i) => i.objectId === object.id);
  const triggerInvPdfUpload = (inv) => { invPdfTarget.current = inv; invPdfRef.current?.click(); };
  const uploadInvoicePdf = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    const inv = invPdfTarget.current;
    if (!file || !inv) return;
    setInvFileBusy(inv.id);
    const path = `${object.id}/invoice_${inv.id}_${Date.now()}_${file.name}`;
    const { error } = await supabase.storage.from("invoices").upload(path, file);
    if (error) { setInvErr(error.message); setInvFileBusy(null); return; }
    await api.setInvoiceFile(inv, path, object.name);
    setInvFileBusy(null);
  };
  const viewInvoicePdf = async (inv) => {
    const { data: s } = await supabase.storage.from("invoices").createSignedUrl(inv.filePath, 300);
    if (s?.signedUrl) setInvViewer({ name: inv.number, url: s.signedUrl, image: /\.(jpe?g|png|gif|webp)$/i.test(inv.filePath) });
  };

  const setIn = (id, patch) => setInputs((p) => ({ ...p, [id]: { ...(p[id] || { from: "", to: "", hours: "" }), ...patch } }));
  const [my, mm] = month.split("-").map(Number);
  const prevM = () => setMonth(mm === 1 ? `${my - 1}-12` : `${my}-${String(mm - 1).padStart(2, "0")}`);
  const nextM = () => setMonth(mm === 12 ? `${my + 1}-01` : `${my}-${String(mm + 1).padStart(2, "0")}`);
  // Neplacené volno za knjigovođu: radni dani u mjesecu bez upisanih sati (po radniku), spojeni u razdoblja od–do
  const exportNeplaceneVolno = () => {
    const month = volnoMonth; const [my, mm] = volnoMonth.split("-").map(Number);
    const CZ_MONTHS = ["leden", "únor", "březen", "duben", "květen", "červen", "červenec", "srpen", "září", "říjen", "listopad", "prosinec"];
    const isHol = (iso) => (object.country === "CZ" ? isCzHoliday(iso) : !!holidayName(iso));
    const lastDay = new Date(my, mm, 0).getDate();
    const days = []; for (let d = 1; d <= lastDay; d++) days.push(`${month}-${String(d).padStart(2, "0")}`);
    const workDays = days.filter((d) => !isWeekend(d) && !isHol(d));
    const monthLogs = data.logs.filter((l) => l.date.startsWith(month));
    const ids = new Set([
      ...data.workers.filter((w) => !w.archived && w.objectId === object.id).map((w) => w.id),
      ...monthLogs.filter((l) => l.objectId === object.id).map((l) => l.workerId),
    ]);
    const people = sortedWorkers(data.workers.filter((w) => ids.has(w.id)));
    const split = (name) => { const t = name.replace(/\s*\(.*?\)\s*/g, " ").trim().split(/\s+/); return { sur: t[0] || "", first: t.slice(1).join(" ") }; };
    const rows = [];
    people.forEach((w) => {
      const mine = monthLogs.filter((l) => l.workerId === w.id);
      const { sur, first } = split(w.name);
      const base = { sur, first, pos: w.position || "", acc: w.bankAccount || "", hours: round2(mine.reduce((s, l) => s + l.hours, 0)) };
      const hasDaily = mine.some((l) => !l.monthly);
      if (!hasDaily) {
        rows.push({ ...base, from: "", to: "", n: "", note: mine.length ? "Hodiny zadány jen za celý měsíc – doplňte dny volna" : "Žádné odpracované hodiny – doplňte" });
        return;
      }
      const worked = new Set(mine.filter((l) => l.hours > 0).map((l) => l.date));
      const missing = workDays.filter((d) => !worked.has(d));
      if (!missing.length) { rows.push({ ...base, from: "", to: "", n: 0, note: "Bez neplaceného volna" }); return; }
      // spoji uzastopne radne dane (vikend/praznik između ne prekida razdoblje)
      const ranges = [];
      missing.forEach((d) => {
        const r = ranges[ranges.length - 1];
        if (r) {
          const between = workDays.slice(workDays.indexOf(r.to) + 1, workDays.indexOf(d));
          if (!between.length) { r.to = d; r.n++; return; }
        }
        ranges.push({ from: d, to: d, n: 1 });
      });
      ranges.forEach((r, i) => rows.push({ ...base, hours: i === 0 ? base.hours : "", acc: i === 0 ? base.acc : "", fd: Number(r.from.slice(8)), td: Number(r.to.slice(8)), from: r.from, to: r.to, n: r.n, note: "" }));
    });
    for (let k = 0; k < 8; k++) rows.push({ sur: "", first: "", pos: "", fd: "", td: "", from: "", to: "", n: "", hours: "", acc: "", note: "" });
    // svátky cijele godine → list "Svátky" za formulu NETWORKDAYS
    const yearDays = []; for (let d = new Date(Date.UTC(my, 0, 1)); d.getUTCFullYear() === my; d.setUTCDate(d.getUTCDate() + 1)) yearDays.push(d.toISOString().slice(0, 10));
    const holidays = yearDays.filter((d) => !isWeekend(d) && isHol(d));
    const holRef = `'Svátky'!$A$5:$A$${4 + Math.max(holidays.length, 1)}`;
    const period = `${CZ_MONTHS[mm - 1]} ${my}`;
    downloadStyledXlsx({
      fileName: `Neplacene_volno_${object.name.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^\w]+/g, "_")}_${month}.xlsx`,
      sheets: [{
        name: "Neplacené volno", title: `Neplacené volno – ${object.name} – ${period}`,
        subtitle: `${data.settings.company_name || "Molnar Rose agency & floristry s.r.o."} · podklad pro mzdy · ${workDays.length} pracovních dní v měsíci`,
        landscape: true, freezeCols: 2, total: { label: "CELKEM", keys: ["n", "hours"] },
        columns: [
          { key: "sur", header: "Příjmení", width: 18, bold: true }, { key: "first", header: "Jméno", width: 16, bold: true },
          { key: "pos", header: "Pozice", width: 16 },
          { key: "fd", header: "Od (den)", width: 9, type: "center", bold: true, fill: () => XC.amberSoft },
          { key: "td", header: "Do (den)", width: 9, type: "center", bold: true, fill: () => XC.amberSoft },
          { key: "from", header: "Neplacené volno od", width: 15, type: "date" },
          { key: "to", header: "Neplacené volno do", width: 15, type: "date" },
          { key: "n", header: "Počet prac. dní", width: 11, type: "int" },
          { key: "hours", header: "Odpracováno hodin", width: 12, type: "hours" },
          { key: "acc", header: "Číslo účtu", width: 24 },
          { key: "note", header: "Poznámka", width: 40, wrap: true, color: (v) => (/doplňte/.test(v || "") ? XC.red : null) },
        ],
        rows,
        emptyText: "V tomto objektu nejsou v tomto měsíci žádní zaměstnanci.",
        note: "Neplacené volno = pracovní dny (po–pá, bez svátků) bez záznamu docházky. Do žlutých sloupců „Od (den)“ a „Do (den)“ stačí napsat číslo dne (např. 15) – datum i počet pracovních dní se doplní samy. Volné řádky dole jsou pro doplnění. Prosím zkontrolujte před odesláním účetní.",
        after: (ws, { first }) => {
          const last = first + rows.length - 1;
          for (let r = first; r <= last; r++) {
            const row = rows[r - first];
            ws.getCell(`F${r}`).value = { formula: `IF(ISNUMBER(D${r}),DATE(${my},${mm},D${r}),"")`, result: row.from ? xDate(row.from) : "" };
            ws.getCell(`G${r}`).value = { formula: `IF(ISNUMBER(E${r}),DATE(${my},${mm},E${r}),"")`, result: row.to ? xDate(row.to) : "" };
            ws.getCell(`H${r}`).value = { formula: `IF(AND(ISNUMBER(F${r}),ISNUMBER(G${r})),NETWORKDAYS(F${r},G${r},${holRef}),"")`, result: typeof row.n === "number" && row.from ? row.n : "" };
            ["F", "G"].forEach((c) => { ws.getCell(`${c}${r}`).numFmt = "dd.mm.yyyy"; });
            ["D", "E", "F", "G", "H"].forEach((c) => { ws.getCell(`${c}${r}`).alignment = { horizontal: "center", vertical: "middle" }; });
          }
          ws.dataValidations.add(`D${first}:E${last}`, { type: "whole", operator: "between", allowBlank: true, formulae: [1, lastDay],
            showErrorMessage: true, errorTitle: "Den v měsíci", error: `Napište jen číslo dne (1–${lastDay}).`, showInputMessage: true, promptTitle: "Den", prompt: `Číslo dne 1–${lastDay}, např. 15` });
        },
      }, {
        name: "Svátky", title: `Státní svátky ${my}`, subtitle: "Nepočítají se do pracovních dní", tabColor: XC.sub,
        columns: [{ key: "d", header: "Datum", width: 14, type: "date" }, { key: "n", header: "Den", width: 14 }],
        rows: holidays.map((d) => ({ d, n: ["neděle", "pondělí", "úterý", "středa", "čtvrtek", "pátek", "sobota"][new Date(d + "T12:00:00").getDay()] })),
      }],
    }).catch((e) => alert("Izvoz nije uspio: " + (e.message || e)));
  };

  const allActiveWorkers = sortedWorkers(data.workers.filter((w) => !w.archived));
  const wq = q.trim().toLowerCase();
  const workers = wq ? allActiveWorkers.filter((w) => w.name.toLowerCase().includes(wq) || (w.position || "").toLowerCase().includes(wq)) : allActiveWorkers;
  const objectPositions = [...new Set(allActiveWorkers.filter((w) => w.objectId === object.id && w.position).map((w) => w.position))].sort((a, b) => a.localeCompare(b, "hr"));
  const myPositionRates = (data.positionBilling || []).filter((pb) => pb.objectId === object.id)
    .sort((a, b) => a.position.localeCompare(b.position, "hr"));
  const editPositionRate = (pb) => {
    setPosRateEditing(pb.position);
    setPosRateForm({ position: pb.position, mode: pb.mode, value: String(pb.mode === "markup" ? pb.markup : pb.rate), currency: pb.currency });
  };
  const resetPositionForm = () => { setPosRateEditing(null); setPosRateForm({ position: "", mode: "fixed", value: "", currency: object.billCur || "EUR" }); };
  const savePositionRate = async () => {
    const pos = posRateForm.position.trim();
    const val = round2(parseNum(posRateForm.value) || 0);
    if (!pos) return;
    if (await api.setPositionRate(object, pos, { mode: posRateForm.mode, value: val, currency: posRateForm.currency })) resetPositionForm();
  };


  const commitDay = (w) => {
    const v = inputs[w.id] || {};
    let hours = parseNum(v.hours), from = v.from || "", to = v.to || "";
    if ((!hours || hours <= 0) && from && to) hours = hoursBetween(from, to);
    if (!hours || hours <= 0) return;
    if (!(from && to)) { from = ""; to = ""; }
    api.addLog({ workerId: w.id, objectId: object.id, date, from, to, hours: round2(hours) }, w.name, object.name);
    setIn(w.id, { from: "", to: "", hours: "" });
  };
  // pri promjeni mjeseca/objekta učitaj postojeće mjesečne sate u polja za uređivanje
  useEffect(() => {
    const vals = {};
    data.workers.filter((w) => !w.archived).forEach((w) => {
      const existing = data.logs.filter((l) => l.workerId === w.id && l.objectId === object.id && l.monthly && monthKey(l.date) === month);
      const sum = round2(existing.reduce((s, l) => s + l.hours, 0));
      vals[w.id] = sum > 0 ? String(sum) : "";
    });
    setMonthVals(vals);
    setMonthSel(new Set());
    setMonthBulkVal("");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [month, object.id]);
  const applyMonthBulkVal = () => {
    if (!monthBulkVal || monthSel.size === 0) return;
    setMonthVals((p) => { const n = { ...p }; monthSel.forEach((id) => { n[id] = monthBulkVal; }); return n; });
  };
  const [showImport, setShowImport] = useState(false);
  // popuni prazna polja sa satima iz prošlog mjeseca (za ovaj objekt) — ništa se ne sprema dok ne klikneš "Spremi"
  const copyPrevMonthVals = () => {
    const pm = prevMonthKey(month);
    let filled = 0;
    const n = { ...monthVals };
    data.workers.filter((w) => !w.archived).forEach((w) => {
      if (n[w.id]) return;
      const sum = round2(data.logs.filter((l) => l.workerId === w.id && l.objectId === object.id && l.monthly && monthKey(l.date) === pm).reduce((s, l) => s + l.hours, 0));
      if (sum > 0) { n[w.id] = String(sum); filled++; }
    });
    if (!filled) { window.alert(tr("Za prošli mjesec na ovom objektu nema mjesečnih sati koje bih mogao prepisati (ili su sva polja već popunjena).")); return; }
    setMonthVals(n);
  };
  const saveMonthBulk = async () => {
    const entries = data.workers.filter((w) => !w.archived).map((w) => ({ workerId: w.id, hours: round2(parseNum(monthVals[w.id]) || 0) }));
    const tooMany = entries.filter((e) => e.hours > MONTH_HOURS_WARN);
    if (tooMany.length && !window.confirm(tr("Više od {0} h u mjesecu ima {1} radnik(a):\n{2}\n\nJe li to točno? Spremiti svejedno?", MONTH_HOURS_WARN, tooMany.length, tooMany.slice(0, 8).map((e) => `• ${data.workers.find((w) => w.id === e.workerId)?.name}: ${fmtH(e.hours)}`).join("\n")))) return;
    setMonthSaving(true);
    await api.setMonthlyHoursBulk(entries, object.id, month, object.name);
    setMonthSaving(false);
  };
  const saveHourEdit = async (mo) => {
    if (!hourEdit) return;
    const hours = round2(parseNum(hourEdit.value) || 0);
    setHourSaving(true);
    await api.setMonthlyHoursBulk([{ workerId: hourEdit.workerId, hours }], object.id, mo, object.name);
    setHourSaving(false);
    setHourEdit(null);
  };
  const commitPosition = (w) => {
    const val = (posEdit[w.id] !== undefined ? posEdit[w.id] : (w.position || "")).trim();
    setPosEdit((p) => { const n = { ...p }; delete n[w.id]; return n; });
    if (val !== (w.position || "")) api.setPosition(w, val);
  };
  const submitTransfer = async () => {
    const ids = [...transferSel].filter((id) => data.workers.find((w) => w.id === id)?.objectId !== object.id);
    if (!ids.length) return;
    setTransferBusy(true);
    if (await api.transferBulk(ids, object.id, object.name, transferDate)) setTransferSel(new Set());
    setTransferBusy(false);
  };

  const entries = data.logs.filter((l) => l.objectId === object.id && (mode === "day" ? l.date === date : monthKey(l.date) === month))
    .sort((a, b) => a.date.localeCompare(b.date));
  const wName = (id) => data.workers.find((x) => x.id === id)?.name || "Obrisan radnik";
  const monthHours = (wid) => round2(data.logs.filter((l) => l.workerId === wid && l.objectId === object.id && monthKey(l.date) === (mode === "day" ? monthKey(date) : month)).reduce((s, l) => s + l.hours, 0));

  return (
    <>
      <button onClick={onBack} style={{ background: "none", border: "none", color: S.green, fontWeight: 700, fontSize: 14.5, padding: "2px 0 12px", cursor: "pointer" }}>{tr("← Natrag")}</button>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
        <div style={{ fontSize: 19, fontWeight: 800 }}>🏨 {object.name}{object.archived && <Tag color={S.sub} bg="#EEF0ED"> {tr("📁 zatvoren")}</Tag>}</div>
        <Tag color={S.blue} bg={S.blueSoft}>{FLAG[object.country || "HR"]} {object.country || tr("HR")}</Tag>
      </div>

      <div style={{ display: "flex", background: "#E9ECE6", borderRadius: 12, padding: 3, marginBottom: 12, gap: 2 }}>
        {[["pregled", "📊 Pregled"], ["sati", "🕒 Sati"], ...(api.admin ? [["fakture", "🧾 Fakture"], ["postavke", "⚙️ Postavke"]] : [])].map(([id, label]) => (
          <button key={id} onClick={() => setSection(id)} style={{ flex: 1, border: "none", borderRadius: 10, padding: "9px 4px", cursor: "pointer", fontSize: 13,
            fontWeight: section === id ? 700 : 500, background: section === id ? "#fff" : "transparent", color: section === id ? S.green : S.sub,
            boxShadow: section === id ? "0 1px 3px rgba(0,0,0,.08)" : "none", whiteSpace: "nowrap" }}>{label}</button>
        ))}
      </div>

      {section === "pregled" && (
        <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA", padding: "10px 14px" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <Btn small kind="ghost" onClick={prevM}>←</Btn>
            <div style={{ fontWeight: 800 }}>{mName(mm - 1)} {my}.</div>
            <Btn small kind="ghost" onClick={nextM}>→</Btn>
          </div>
        </Card>
      )}

      {api.admin && (() => {
        const mk = section === "sati" && mode === "day" ? monthKey(date) : month;
        const objLogs = data.logs.filter((l) => l.objectId === object.id && monthKey(l.date) === mk);
        const objBrigade = (data.brigadeCosts || []).filter((bg) => bg.objectId === object.id && monthKey(bg.date) === mk);
        const rv = objRevenue(data, object, objLogs, objBrigade);
        const bc = object.billCur || "EUR";
        let costE = 0, costK = 0;
        objLogs.forEach((l) => {
          const wk = data.workers.find((x) => x.id === l.workerId);
          const v = l.hours * rateFor(data, wk, l.date) + extraPayForLog(data, wk, l);
          if (wCur(wk) === "CZK") costK += v; else costE += v;
        });
        costE = round2(costE); costK = round2(costK);
        const oPays = data.payments.filter((pp) => !pp.workerId && pp.objectId === object.id && monthKey(pp.date) === mk);
        const brigadeCostE = round2(objBrigade.filter((bg) => (bg.currency || "CZK") === "EUR").reduce((s, bg) => s + bg.amount, 0));
        const brigadeCostK = round2(objBrigade.filter((bg) => bg.currency === "CZK").reduce((s, bg) => s + bg.amount, 0));
        const extraCost = round2(oPays.filter((pp) => (pp.currency || "EUR") === "EUR").reduce((s, pp) => s + pp.amount, 0));
        const extraCostKc = round2(oPays.filter((pp) => pp.currency === "CZK").reduce((s, pp) => s + pp.amount, 0));
        return (
          <>
          {section === "sati" && <BrigadeCard object={object} data={data} api={api} monthBrigade={objBrigade} defaultDate={mode === "day" ? date : todayISO()} />}
          {section === "pregled" && (
          <Card style={{ background: S.amberSoft, borderColor: "#EBD9B4" }}>
            <div style={{ fontWeight: 800, fontSize: 15, marginBottom: 10 }}>{tr("💶 Naplata i dobit —")}{" "}{mName(mm - 1)} {my}.</div>
            {(() => {
              const revE = rv.revenue, revK = rv.revenueKc;
              const profitE = round2(revE - costE - extraCost - brigadeCostE);
              const profitK = round2(revK - costK - extraCostKc - brigadeCostK);
              return (
                <div className="num" style={{ fontSize: 14 }}>
                  {rv.byPosition.length === 0 && (
                    <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Naplata (0 h)")}</span><b>{money(0, bc)}</b></div>
                  )}
                  {rv.byPosition.map((row, i) => (
                    <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}>
                      <span>{tr("Naplata")}{row.position ? " · " + row.position : ""} ({fmtH(row.hours)} × {money(row.rate, row.cur)})</span><b>{money(row.amount, row.cur)}</b>
                    </div>
                  ))}
                  {costE > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Plaća radnika (€)")}</span><b>−{eur(costE)}</b></div>}
                  {costK > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Plaća radnika (Kč)")}</span><b>−{czk(costK)}</b></div>}
                  {extraCost > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Troškovi objekta (€)")}</span><b>−{eur(extraCost)}</b></div>}
                  {extraCostKc > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Troškovi objekta (Kč)")}</span><b>−{czk(extraCostKc)}</b></div>}
                  {brigadeCostE > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Brigada — isplaćeno cash (€)")}</span><b>−{eur(brigadeCostE)}</b></div>}
                  {brigadeCostK > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Brigada — isplaćeno cash (Kč)")}</span><b>−{czk(brigadeCostK)}</b></div>}
                  <div style={{ borderTop: `1px dashed #EBD9B4`, paddingTop: 5, marginTop: 2 }}>
                    {(profitE !== 0 || profitK === 0) && (
                      <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 800 }}>
                        <span>{tr("Dobit ovaj mjesec (€)")}</span><span style={{ color: profitE >= 0 ? S.green : S.red }}>{eur(profitE)}</span>
                      </div>
                    )}
                    {profitK !== 0 && (
                      <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 800 }}>
                        <span>{tr("Dobit ovaj mjesec (Kč)")}</span><span style={{ color: profitK >= 0 ? S.green : S.red }}>{czk(profitK)}</span>
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}
            {object.monthlyBudget > 0 && (() => { const spent = bc === "CZK" ? costK + extraCostKc + brigadeCostK : costE + extraCost + brigadeCostE; const pct = round2((spent / object.monthlyBudget) * 100); const over = spent > object.monthlyBudget;
              return <div className="num" style={{ fontSize: 12.5, marginTop: 8, fontWeight: 600, color: over ? S.red : pct >= 80 ? S.amber : S.sub }}>{over || pct >= 80 ? "⚠️ " : ""}{tr("Proračun: potrošeno")}{" "}{money(round2(spent), bc)} {tr("od")}{" "}{money(object.monthlyBudget, bc)} ({pct}%)</div>; })()}
          </Card>
          )}
          {section === "postavke" && (<>
          <Card>
            <div style={{ fontWeight: 800, fontSize: 15, marginBottom: 10 }}>{tr("💶 Naplata objekta")}</div>
            <Field label={tr("Objekt plaća ({0} / sat)", bc === "CZK" ? "Kč" : "€")}>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input inputMode="decimal" value={rateEdit} onChange={(e) => setRateEdit(e.target.value)} placeholder={bc === "CZK" ? tr("npr. 250") : tr("npr. 10.00")} />
                <CurChips small value={billCur} onChange={setBillCur} />
                <Btn small onClick={() => api.setBillRate(object, parseNum(rateEdit) || 0, billCur)}>{tr("Spremi")}</Btn>
              </div>
              <div style={{ fontSize: 11.5, color: S.sub, marginTop: 4 }}>{tr("Ovo je zadana (default) cijena za sve pozicije koje nemaju svoju cijenu ispod.")}</div>
            </Field>

            <Field label={tr("Naziv kupca na PDF fakturi (za automatski uvoz) — nije obavezno")}>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input value={custNameEdit} onChange={(e) => setCustNameEdit(e.target.value)} placeholder={tr("npr. F&B Solutions s.r.o.")} />
                <Btn small onClick={() => api.setInvoiceCustomerName(object, custNameEdit.trim())}>{tr("Spremi")}</Btn>
              </div>
              <div style={{ fontSize: 11.5, color: S.sub, marginTop: 4 }}>
                {tr("Mora točno odgovarati imenu \"Odběratel\" na fakturi — Objekti tab → \"📥 Uvoz faktura\" će onda ovaj objekt sam prepoznati.")}
              </div>
            </Field>

            <Field label={tr("Mjesečni proračun troška ({0}) — nije obavezno", bc === "CZK" ? "Kč" : "€")}>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input inputMode="decimal" value={budgetEdit} onChange={(e) => setBudgetEdit(e.target.value)} placeholder={tr("npr. 5000")} />
                <Btn small onClick={() => api.setBudget(object, parseNum(budgetEdit) || 0, bc)}>{tr("Spremi")}</Btn>
              </div>
              {object.monthlyBudget > 0 && (() => {
                const spent = bc === "CZK" ? costK + extraCostKc + brigadeCostK : costE + extraCost + brigadeCostE;
                const pct = round2((spent / object.monthlyBudget) * 100);
                const over = spent > object.monthlyBudget;
                return (
                  <div style={{ fontSize: 12.5, marginTop: 4, fontWeight: 600, color: over ? S.red : pct >= 80 ? S.amber : S.sub }}>
                    {over ? "⚠️ " : pct >= 80 ? "⚠️ " : ""}{tr("Potrošeno")}{" "}{money(round2(spent), bc)} {tr("od")}{" "}{money(object.monthlyBudget, bc)} ({pct}%){over ? tr(" — PREKORAČENO") : ""}
                  </div>
                );
              })()}
            </Field>

          </Card>
          <Card>
            <div style={{ fontWeight: 800, fontSize: 15, marginBottom: 10 }}>{tr("🏷 Cijene po poziciji")}</div>
            <div style={{ marginTop: 0 }}>
              <div style={{ fontSize: 12, color: S.sub, marginBottom: 8 }}>{tr("Opcionalno: drugačija cijena za pojedinu poziciju.")}</div>
              {myPositionRates.length > 0 && (
                <div style={{ marginBottom: 8 }}>
                  {myPositionRates.map((pb) => (
                    <div key={pb.position} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", borderBottom: `1px solid #EBD9B4`, fontSize: 13 }}>
                      <span style={{ flex: 1, fontWeight: 700 }}>{pb.position}</span>
                      <span className="num" style={{ color: S.sub }}>
                        {pb.mode === "markup" ? <>{tr("satnica radnika +")}{" "}{money(pb.markup, pb.currency)}</> : <>{money(pb.rate, pb.currency)}/h</>}
                      </span>
                      <button onClick={() => editPositionRate(pb)} style={{ background: "none", border: "none", color: S.blue, fontSize: 14, cursor: "pointer", padding: 4 }}>✎</button>
                      <button onClick={() => api.delPositionRate(object, pb.position)} style={{ background: "none", border: "none", color: S.red, fontSize: 15, cursor: "pointer", padding: 4 }}>✕</button>
                    </div>
                  ))}
                </div>
              )}
              <div style={{ display: "flex", gap: 8, alignItems: "flex-end", flexWrap: "wrap" }}>
                <div style={{ flex: "1 1 130px" }}>
                  <Field label={tr("Pozicija")}>
                    <input list="obj-positions" value={posRateForm.position} disabled={!!posRateEditing}
                      onChange={(e) => setPosRateForm({ ...posRateForm, position: e.target.value })} placeholder={tr("npr. HSK")} />
                    <datalist id="obj-positions">{objectPositions.map((p) => <option key={p} value={p} />)}</datalist>
                  </Field>
                </div>
                <div style={{ flex: "1 1 150px" }}>
                  <Field label={tr("Način")}>
                    <select value={posRateForm.mode} onChange={(e) => setPosRateForm({ ...posRateForm, mode: e.target.value })}>
                      <option value="fixed">{tr("Fiksna cijena / sat")}</option>
                      <option value="markup">{tr("Satnica radnika + toliko")}</option>
                    </select>
                  </Field>
                </div>
                <div style={{ flex: "1 1 110px" }}>
                  <Field label={posRateForm.mode === "markup" ? tr("+ po satu") : tr("cijena / sat")}>
                    <input inputMode="decimal" value={posRateForm.value} onChange={(e) => setPosRateForm({ ...posRateForm, value: e.target.value })} placeholder="0.00" />
                  </Field>
                </div>
                <CurChips small value={posRateForm.currency} onChange={(v) => setPosRateForm({ ...posRateForm, currency: v })} />
                <div style={{ marginBottom: 10, display: "flex", gap: 6 }}>
                  <Btn small onClick={savePositionRate}>{posRateEditing ? tr("Spremi") : tr("Dodaj")}</Btn>
                  {posRateEditing && <Btn small kind="ghost" onClick={resetPositionForm}>{tr("Odustani")}</Btn>}
                </div>
              </div>
            </div>

          </Card>
          <Card>
            <div style={{ fontWeight: 800, fontSize: 15, marginBottom: 10 }}>{tr("🔑 Pristup zaposlenika")}</div>
            <div style={{ marginTop: 0 }}>
              <div style={{ fontSize: 12, color: S.sub, marginBottom: 6 }}>{tr("Koji zaposlenici vide ovaj objekt u aplikaciji.")}</div>
              {(data.profiles || []).filter((p) => p.role !== "admin").length === 0 ? (
                <div style={{ fontSize: 12.5, color: S.sub }}>{tr("Još nema zaposlenika.")}</div>
              ) : [...(data.profiles || [])].filter((p) => p.role !== "admin").sort((a, b) => (a.name || "").localeCompare(b.name || "", "hr")).map((p) => {
                const on = (data.objectMembers || []).some((m) => m.object_id === object.id && m.user_id === p.id);
                return (
                  <label key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", fontSize: 13.5, cursor: "pointer" }}>
                    <input type="checkbox" checked={on} onChange={(e) => api.toggleMember(object, p, e.target.checked)} style={{ width: 18, height: 18 }} />
                    {p.name || tr("(bez imena)")}
                  </label>
                );
              })}
            </div>
          </Card>
          <Card>
            <div style={{ fontWeight: 800, fontSize: 15, marginBottom: 10 }}>{tr("🏳️ Država i status")}</div>
            <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
          {["HR", "CZ"].map((c) => (
            <button key={c} onClick={() => (object.country || "HR") !== c && api.setObjectCountry(object, c)} style={{
              padding: "5px 11px", borderRadius: 999, fontWeight: 700, fontSize: 12.5, cursor: "pointer",
              background: (object.country || "HR") === c ? S.blue : "#fff", color: (object.country || "HR") === c ? "#fff" : S.sub,
              border: `1px solid ${(object.country || "HR") === c ? S.blue : S.line}` }}>
              {FLAG[c]} {c}
            </button>
          ))}
          {api.admin && (
            <Btn small kind="ghost" onClick={() => api.archiveObject(object, !object.archived)} style={{ color: object.archived ? S.green : S.amber, fontWeight: 700 }}>
              {object.archived ? tr("↩ Otvori objekt") : tr("📁 Zatvori objekt")}
            </Btn>
          )}
            </div>
          </Card>
          </>)}
          </>
        );
      })()}

      {section === "pregled" && (() => {
        const mk2 = month;
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
            <div style={{ fontWeight: 800, fontSize: 15, marginBottom: 6 }}>{tr("👷 Radnici i sati —")}{" "}{mName(mm - 1)} {my}.</div>
            {rows2.map((r) => (
              <div key={r.w.id} className="num" style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13.5 }}>
                <span onClick={() => onOpenWorker && onOpenWorker(r.w.id)} style={{ fontWeight: 600, cursor: onOpenWorker ? "pointer" : "default", color: onOpenWorker ? S.blue : "inherit" }}>
                  {r.w.name}{r.w.objectId === object.id ? <span style={{ color: S.blue, fontSize: 11, fontWeight: 700 }}> ★</span> : ""}
                </span>
                {hourEdit && hourEdit.workerId === r.w.id ? (
                  <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <input inputMode="decimal" autoFocus value={hourEdit.value}
                      onChange={(e) => setHourEdit({ workerId: r.w.id, value: e.target.value })}
                      onKeyDown={(e) => e.key === "Enter" && saveHourEdit(mk2)}
                      style={{ width: 72, padding: "4px 6px", fontSize: 13 }} />
                    <button onClick={() => saveHourEdit(mk2)} disabled={hourSaving} style={{ background: "none", border: "none", color: S.green, fontWeight: 800, cursor: "pointer", fontSize: 17, padding: 2 }}>✓</button>
                    <button onClick={() => setHourEdit(null)} style={{ background: "none", border: "none", color: S.red, fontWeight: 800, cursor: "pointer", fontSize: 15, padding: 2 }}>✕</button>
                  </span>
                ) : (
                  <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                    {isHoursSuspicious(r.h) && <span title={tr("Više od {0}h u mjesecu — provjeri unos", MONTH_HOURS_WARN)}>⚠️</span>}
                    <span onClick={() => setHourEdit({ workerId: r.w.id, value: String(r.h) })} style={{ fontWeight: 700, cursor: "pointer", color: isHoursSuspicious(r.h) ? S.red : S.blue }}>
                      {fmtH(r.h)} <span style={{ fontSize: 12 }}>✎</span>
                    </span>
                  </span>
                )}
              </div>
            ))}
            <div className="num" style={{ display: "flex", justifyContent: "space-between", paddingTop: 7, fontWeight: 800 }}>
              <span>{tr("Ukupno")}</span><span>{fmtH(tot)}</span>
            </div>
          </Card>
        );
      })()}

      {api.admin && section === "fakture" && (
        <Card>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
            <div style={{ fontWeight: 700 }}>{tr("🧾 Fakture i dokumenti objekta")}</div>
            <div style={{ display: "flex", gap: 6 }}>
              <Btn small onClick={() => (showGen ? setShowGen(false) : openGenerator())}>{showGen ? tr("Zatvori") : tr("⚡ Generiraj fakturu")}</Btn>
              <Btn small kind="ghost" onClick={() => invFileRef.current?.click()}>{tr("+ Dodaj datoteku")}</Btn>
            </div>
            <input ref={invFileRef} type="file" accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.txt" multiple onChange={uploadInvoice} style={{ display: "none" }} />
            <input ref={invPdfRef} type="file" accept="application/pdf,image/*" onChange={uploadInvoicePdf} style={{ display: "none" }} />
          </div>

          {showGen && (
            <div style={{ background: S.greenSoft, borderRadius: 10, padding: "12px 14px", marginBottom: 14 }}>
              <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 10 }}>
                {tr("Podaci su povučeni iz upisanih sati i naplate objekta — po potrebi ih slobodno promijeni prije nego stisneš Generiraj.")}
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <div style={{ flex: 1 }}><Field label={tr("Mjesec")}><input type="month" value={genMonth} onChange={(e) => pickGenMonth(e.target.value)} /></Field></div>
                <div style={{ flex: 1 }}><Field label={tr("Broj računa")}><input value={genNumber} onChange={(e) => setGenNumber(e.target.value)} placeholder={tr("npr. ADR-2026-08")} /></Field></div>
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <div style={{ flex: 1 }}><Field label={tr("Sati")}><input inputMode="decimal" value={genHours} onChange={(e) => setGenHours(e.target.value)} /></Field></div>
                <div style={{ flex: 1 }}>
                  <Field label={tr("Cijena po satu ({0})", genCur === "CZK" ? "Kč" : "€")}>
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <input inputMode="decimal" value={genRate} onChange={(e) => setGenRate(e.target.value)} />
                      <CurChips small value={genCur} onChange={setGenCur} />
                    </div>
                  </Field>
                </div>
              </div>
              <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, fontSize: 13.5, cursor: "pointer" }}>
                <input type="checkbox" checked={autoAmount} onChange={(e) => { setAutoAmount(e.target.checked); if (!e.target.checked) setGenAmountManual(String(genAuto)); }} style={{ width: 18, height: 18 }} />
                {tr("Iznos automatski (sati × cijena)")}
              </label>
              <Field label={tr("Iznos za naplatu ({0})", genCur === "CZK" ? "Kč" : "€")}>
                {autoAmount
                  ? <div style={{ background: "#fff", border: `1px solid ${S.line}`, borderRadius: 10, padding: "11px 12px", fontWeight: 700 }}>{money(genAuto, genCur)}</div>
                  : <input inputMode="decimal" value={genAmountManual} onChange={(e) => setGenAmountManual(e.target.value)} />}
              </Field>
              <Field label={tr("Napomena na fakturi (nije obavezno)")}><input value={genNote} onChange={(e) => setGenNote(e.target.value)} placeholder={tr("npr. uključuje prekovremene")} /></Field>
              <Field label={tr("PDV na fakturi (iznos koji vidiš na PDF-u, nije obavezno)")}>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <input inputMode="decimal" value={genVat} onChange={(e) => setGenVat(e.target.value)} placeholder="0.00" />
                  <CurChips small value={genVatCur} onChange={setGenVatCur} />
                </div>
              </Field>
              <Btn onClick={submitInvoice} style={{ width: "100%" }}>{tr("🧾 Generiraj i spremi fakturu")}</Btn>
            </div>
          )}

          {invErr && <div style={{ color: S.red, fontSize: 13, marginBottom: 8 }}>{invErr}</div>}

          {myInvoices.length > 0 && (
            <div style={{ marginBottom: 14 }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, color: S.sub, marginBottom: 6 }}>{tr("Generirane fakture")}</div>
              {myInvoices.map((inv) => (
                <div key={inv.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderBottom: `1px solid ${S.line}`, flexWrap: "wrap" }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 600, fontSize: 13.5 }}>{inv.number}</div>
                    <div className="num" style={{ fontSize: 12, color: S.sub }}>
                      {fmtDate(inv.issueDate)} · {fmtH(inv.hours)} · {money(inv.amount, inv.currency)}
                      {inv.vatAmount > 0 && <> {tr("· PDV")}{" "}{money(inv.vatAmount, inv.vatCurrency)}</>}
                    </div>
                  </div>
                  <Btn small kind="ghost" onClick={() => printGeneratedInvoice(inv, object, data.settings || {})}>🖨</Btn>
                  {inv.filePath ? (
                    <>
                      <Btn small kind="ghost" onClick={() => viewInvoicePdf(inv)}>{tr("📄 PDF")}</Btn>
                      <button onClick={() => api.delInvoiceFile(inv, object.name)} title={tr("Ukloni PDF")} style={{ background: "none", border: "none", color: S.red, fontSize: 13, cursor: "pointer", padding: 4 }}>{tr("PDF ✕")}</button>
                    </>
                  ) : (
                    <Btn small kind="ghost" onClick={() => triggerInvPdfUpload(inv)} disabled={invFileBusy === inv.id}>
                      {invFileBusy === inv.id ? tr("Učitavam…") : tr("+ PDF")}
                    </Btn>
                  )}
                  <button onClick={() => api.delObjectInvoice(inv, object.name)} style={{ background: "none", border: "none", color: S.red, fontSize: 16, cursor: "pointer", padding: 4 }}>✕</button>
                </div>
              ))}
            </div>
          )}

          {invoices === null ? <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Učitavam…")}</div>
            : invoices.length === 0 ? <div style={{ color: S.sub, fontSize: 13.5 }}>{tr("Slikaj ili dodaj PDF fakture koje stižu od hotela — ostat će spremljene ovdje.")}</div>
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

      {section === "pregled" && (
            <Card style={{ padding: "10px 14px", display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <div style={{ flex: 1, minWidth: 180 }}>
                <div style={{ fontWeight: 700, fontSize: 14 }}>{tr("📄 Neplacené volno za knjigovođu")}</div>
                <div style={{ fontSize: 12, color: S.sub }}>{tr("Tko, od kada do kada, broj računa — za knjigovođu")}</div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                <Btn small kind="ghost" onClick={() => setVolnoMonth(prevMonthKey(volnoMonth))}>←</Btn>
                <div style={{ fontWeight: 700, fontSize: 13.5, minWidth: 104, textAlign: "center" }}>{mName(Number(volnoMonth.slice(5)) - 1)} {volnoMonth.slice(0, 4)}.</div>
                <Btn small kind="ghost" onClick={() => { const [yy, m2] = volnoMonth.split("-").map(Number); setVolnoMonth(m2 === 12 ? `${yy + 1}-01` : `${yy}-${String(m2 + 1).padStart(2, "0")}`); }}>→</Btn>
              </div>
              <Btn small kind="excel" onClick={exportNeplaceneVolno}>{tr("⬇ Excel")}</Btn>
            </Card>
      )}

      {section === "sati" && (<>
      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
        {[["day", "Unos po danu"], ["month", "Sati za cijeli mjesec"]].map(([id, label]) => (
          <button key={id} onClick={() => setMode(id)} style={{ flex: 1, padding: "10px 8px", borderRadius: 10, fontWeight: 700, fontSize: 13.5, cursor: "pointer",
            background: mode === id ? S.blue : "#fff", color: mode === id ? "#fff" : S.sub, border: `1px solid ${mode === id ? S.blue : S.line}` }}>{label}</button>
        ))}
      </div>

      {mode === "day" ? (
        <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA", padding: "10px 14px" }}>
          <Field label={tr("Datum")}><input type="date" value={date} onChange={(e) => setDate(e.target.value)} /></Field>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
            <DayBadge iso={date} />
            <Btn small kind="ghost" onClick={() => api.copyDay(object, addDaysISO(date, -1), date)} style={{ marginLeft: "auto" }}>
              {tr("📋 Kopiraj jučerašnji dan")}
            </Btn>
          </div>
        </Card>
      ) : (
        <>
          <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA", padding: "10px 14px" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <Btn small kind="ghost" onClick={prevM}>←</Btn>
              <div style={{ fontWeight: 800 }}>{mName(mm - 1)} {my}.</div>
              <Btn small kind="ghost" onClick={nextM}>→</Btn>
            </div>
          </Card>
          <Card style={{ padding: "12px 14px" }}>
            <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, cursor: "pointer" }}>
              <input type="checkbox" checked={workers.length > 0 && monthSel.size === workers.length}
                onChange={(e) => setMonthSel(e.target.checked ? new Set(workers.map((w) => w.id)) : new Set())}
                style={{ width: 18, height: 18 }} />
              <span style={{ fontSize: 13.5, fontWeight: 700 }}>{tr("Odaberi sve (")}{monthSel.size}/{workers.length})</span>
            </label>
            <div style={{ display: "flex", gap: 8 }}>
              <input inputMode="decimal" placeholder={tr("sati")} value={monthBulkVal} onChange={(e) => setMonthBulkVal(e.target.value)} style={{ flex: 1 }} />
              <Btn small kind="ghost" onClick={applyMonthBulkVal} disabled={monthSel.size === 0 || !monthBulkVal}>
                {tr("Postavi odabranima")}{" "}{monthSel.size > 0 ? `(${monthSel.size})` : ""}
              </Btn>
            </div>
            <Btn small kind="ghost" onClick={copyPrevMonthVals} style={{ width: "100%", marginTop: 8 }}>
              {tr("↺ Prepiši sate iz prošlog mjeseca (samo prazna polja)")}
            </Btn>
            <Btn small kind="excel" onClick={() => setShowImport(!showImport)} style={{ width: "100%", marginTop: 8 }}>
              {tr("📥 Uvezi sate iz Excela hotela")}
            </Btn>
            <div style={{ fontSize: 12, color: S.sub, marginTop: 6 }}>
              {tr("Označi radnike kvačicom i upiši sate pa \"Postavi odabranima\" — postavlja se svima odjednom (za jednoga, za sve ili samo za odabrane). Svaki broj možeš i ručno urediti u retku ispod. Za kraj klikni \"Spremi sate za mjesec\".")}
            </div>
          </Card>
          {showImport && (
            <HoursImport data={data} object={object} month={month} onClose={() => setShowImport(false)}
              onSave={(entries) => api.setMonthlyHoursBulk(entries, object.id, month, object.name + " · uvoz iz Excela")} />
          )}
        </>
      )}

      <Card style={{ background: S.blueSoft, borderColor: "#CBDCEA" }}>
        <div onClick={() => setShowTransfer(!showTransfer)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
          <div style={{ fontWeight: 700, color: S.blue }}>{tr("👥 Grupni transfer na ovaj objekt")}{transferSel.size > 0 ? ` (${transferSel.size})` : ""}</div>
          <span style={{ color: S.blue, fontWeight: 700 }}>{showTransfer ? "▲" : "▼"}</span>
        </div>
        {showTransfer && (
          <div style={{ marginTop: 10 }}>
            <div style={{ fontSize: 12, color: S.sub, marginBottom: 8 }}>{tr("Označi radnike koje seliš na ovaj objekt (npr. za novu sezonu) i klikni jednom \"Prebaci odabrane\".")}</div>
            <Field label={tr("Prebačeni od datuma")}><input type="date" value={transferDate} onChange={(e) => setTransferDate(e.target.value)} /></Field>
            <div style={{ maxHeight: 260, overflowY: "auto", marginBottom: 10 }}>
              {allActiveWorkers.filter((w) => w.objectId !== object.id).map((w) => (
                <label key={w.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px solid #CBDCEA", fontSize: 13.5, cursor: "pointer" }}>
                  <input type="checkbox" checked={transferSel.has(w.id)}
                    onChange={(e) => setTransferSel((p) => { const n = new Set(p); e.target.checked ? n.add(w.id) : n.delete(w.id); return n; })}
                    style={{ width: 17, height: 17 }} />
                  <span style={{ flex: 1 }}>{w.name}</span>
                  <span style={{ fontSize: 12, color: S.sub }}>{data.objects.find((o) => o.id === w.objectId)?.name || tr("bez objekta")}</span>
                </label>
              ))}
              {allActiveWorkers.filter((w) => w.objectId !== object.id).length === 0 && (
                <div style={{ fontSize: 12.5, color: S.sub }}>{tr("Svi aktivni radnici su već na ovom objektu.")}</div>
              )}
            </div>
            <Btn small onClick={submitTransfer} disabled={transferBusy || transferSel.size === 0} style={{ width: "100%" }}>
              {transferBusy ? tr("Prebacujem…") : tr("✓ Prebaci odabrane ({0})", transferSel.size)}
            </Btn>
          </div>
        )}
      </Card>

      {allActiveWorkers.length > 5 && (
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder={tr("🔍 Traži radnika po imenu ili poziciji…")} style={{ marginBottom: 10 }} />
      )}
      {q.trim() && workers.length === 0 && <Empty text={tr("Nema rezultata.")} />}

      {workers.map((w) => {
        const v = inputs[w.id] || { from: "", to: "", hours: "" };
        return (
          <Card key={w.id} style={{ padding: "12px 14px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div onClick={() => onOpenWorker && onOpenWorker(w.id)} style={{ cursor: onOpenWorker ? "pointer" : "default" }}>
                <span style={{ fontWeight: 700, fontSize: 15 }}>{w.name}</span>
                {onOpenWorker && <span style={{ color: S.blue, fontWeight: 700 }}> ›</span>}
                {w.objectId === object.id && <span style={{ marginLeft: 6, fontSize: 11, color: S.blue, fontWeight: 700 }}>{tr("★ ovaj objekt")}</span>}
              </div>
              <div className="num" style={{ fontSize: 12.5, color: S.sub }}>{fmtH(monthHours(w.id))} {mode === "day" ? tr("ovdje ovaj mj.") : tr("ovdje taj mjesec")}</div>
            </div>
            <input
              value={posEdit[w.id] !== undefined ? posEdit[w.id] : (w.position || "")}
              onChange={(e) => setPosEdit((p) => ({ ...p, [w.id]: e.target.value }))}
              onBlur={() => commitPosition(w)}
              onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
              placeholder={tr("pozicija — npr. HSK, kuhinja, bar…")}
              style={{ fontSize: 12.5, padding: "5px 9px", width: "100%", marginBottom: 8, borderRadius: 8, border: `1px solid ${S.line}`, color: w.position ? S.blue : S.sub, fontWeight: w.position ? 700 : 400 }}
            />
            {mode === "day" ? (
              <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <input type="time" value={v.from} onChange={(e) => setIn(w.id, { from: e.target.value })} style={{ flex: 1.2, padding: "9px 6px" }} />
                <span style={{ color: S.sub }}>–</span>
                <input type="time" value={v.to} onChange={(e) => setIn(w.id, { to: e.target.value })} style={{ flex: 1.2, padding: "9px 6px" }} />
                <input inputMode="decimal" placeholder={tr("ili sati")} value={v.hours} onChange={(e) => setIn(w.id, { hours: e.target.value })} style={{ flex: 1, padding: "9px 8px" }} />
                <Btn small onClick={() => commitDay(w)}>{tr("Upiši")}</Btn>
              </div>
            ) : (
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input type="checkbox" checked={monthSel.has(w.id)}
                  onChange={(e) => setMonthSel((p) => { const n = new Set(p); e.target.checked ? n.add(w.id) : n.delete(w.id); return n; })}
                  style={{ width: 18, height: 18, flex: "0 0 auto" }} />
                <input inputMode="decimal" placeholder={tr("ukupno sati za mjesec")} value={monthVals[w.id] || ""}
                  onChange={(e) => setMonthVals((p) => ({ ...p, [w.id]: e.target.value }))} />
              </div>
            )}
          </Card>
        );
      })}

      {mode === "month" && workers.length > 0 && (
        <Btn onClick={saveMonthBulk} disabled={monthSaving} style={{ width: "100%", marginBottom: 12 }}>
          {monthSaving ? tr("Spremam…") : tr("💾 Spremi sate za mjesec")}
        </Btn>
      )}

      {entries.length > 0 && (
        <Card>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("Upisi —")}{" "}{mode === "day" ? fmtDate(date) : `${mName(mm - 1)} ${my}.`}</div>
          {entries.map((l) => (
            <div key={l.id} style={{ borderBottom: `1px solid ${S.line}` }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0" }}>
                <div style={{ flex: 1 }}>
                  <span onClick={() => onOpenWorker && onOpenWorker(l.workerId)} style={{ fontWeight: 600, fontSize: 14, cursor: onOpenWorker ? "pointer" : "default", color: onOpenWorker ? S.blue : "inherit" }}>{wName(l.workerId)}</span>
                  <div className="num" style={{ fontSize: 12.5, color: S.sub }}>{mode === "month" ? fmtDate(l.date) + " · " : ""}<LogSpanLabel l={l} /></div>
                </div>
                <div className="num" style={{ fontWeight: 700, color: !l.monthly && isDayHoursSuspicious(l.hours) ? S.red : "inherit" }}>
                  {!l.monthly && isDayHoursSuspicious(l.hours) && "⚠️ "}{fmtH(l.hours)}
                </div>
                {(api.admin || l.createdBy === api.uid()) ? (
                  <>
                    <button onClick={() => setEditLog(editLog === l.id ? null : l.id)} style={{ background: "none", border: "none", color: S.blue, fontSize: 15, cursor: "pointer", padding: 4 }}>✎</button>
                    <button onClick={() => api.delLog(l, wName(l.workerId))} style={{ background: "none", border: "none", color: S.red, fontSize: 16, cursor: "pointer", padding: 4 }}>✕</button>
                  </>
                ) : <span title={tr("Upisao kolega")} style={{ fontSize: 13, color: S.sub, padding: 4 }}>👤</span>}
              </div>
              {editLog === l.id && (
                <LogEditor log={l} onCancel={() => setEditLog(null)}
                  onSave={async (patch) => { if (await api.updLog(l, patch, wName(l.workerId))) setEditLog(null); }} />
              )}
            </div>
          ))}
          <div className="num" style={{ display: "flex", justifyContent: "space-between", fontWeight: 800, paddingTop: 8 }}>
            <span>{tr("Ukupno")}</span><span>{fmtH(round2(entries.reduce((s, l) => s + l.hours, 0)))}</span>
          </div>
        </Card>
      )}

      </>)}

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
            <a href={invViewer.url} target="_blank" rel="noreferrer" style={{ background: S.green, color: "#fff", borderRadius: 10, padding: "9px 14px", fontWeight: 700, textDecoration: "none", fontSize: 13.5 }}>{tr("⬇ Otvori / preuzmi")}</a>
            <Btn small kind="ghost" onClick={(e) => { e.stopPropagation(); delInvoice({ name: invViewer.name }); }}>{tr("Obriši")}</Btn>
            <Btn small onClick={() => setInvViewer(null)}>{tr("Zatvori")}</Btn>
          </div>
        </div>
      )}
    </>
  );
}

/* ================================================================== */
/*  SATI + podsjetnik na neupisane dane                                */
/* ================================================================== */
function HoursTab({ data, api, onOpenWorker }) {
  const [mode, setMode] = useState("day"); // day | month
  const [form, setForm] = useState({ workerId: "", objectId: "", date: todayISO(), from: "07:00", to: "15:00", note: "" });
  const [monthBulk, setMonthBulk] = useState({ month: curMonth(), objectId: "" });
  const [monthHours, setMonthHours] = useState({}); // workerId -> string
  const [monthBusy, setMonthBusy] = useState(false);
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
  const monthWorkers = workersAtObject(data, monthBulk.objectId);
  const monthEntryCount = Object.values(monthHours).filter((v) => (parseNum(v) || 0) > 0).length;
  const submitMonthBulk = async () => {
    const entries = Object.entries(monthHours)
      .map(([workerId, v]) => ({ workerId, hours: round2(parseNum(v) || 0) }))
      .filter((e) => e.hours > 0)
      .map((e) => ({ ...e, objectId: monthBulk.objectId || data.workers.find((w) => w.id === e.workerId)?.objectId || "" }));
    if (!entries.length) return;
    setMonthBusy(true);
    if (await api.addLogsBulk(entries, monthBulk.month)) setMonthHours({});
    setMonthBusy(false);
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
            <b style={{ color: S.amber }}>{tr("⏰ Dani bez ijednog upisa ovaj mjesec:")}</b>{" "}
            <span className="num">{missing.slice(0, 8).map((d) => fmtDate(d).slice(0, -1)).join(", ")}{missing.length > 8 ? ` +${missing.length - 8}` : ""}</span>
          </div>
        </Card>
      )}

      {data.workers.filter((x) => !x.archived).length === 0 ? <Empty text={tr("Prvo dodaj radnika.")} /> : (
        <>
          <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
            {[["day", "Po danu"], ["month", "Cijeli mjesec"]].map(([id, label]) => (
              <button key={id} onClick={() => setMode(id)} style={{ flex: 1, padding: "10px 8px", borderRadius: 10, fontWeight: 700, fontSize: 13.5, cursor: "pointer",
                background: mode === id ? S.blue : "#fff", color: mode === id ? "#fff" : S.sub, border: `1px solid ${mode === id ? S.blue : S.line}` }}>{label}</button>
            ))}
          </div>
          {mode === "day" ? (
            <Card>
              <Field label={api.t("object")}><ObjectSelect data={data} api={api} value={form.objectId} onChange={pickObject} /></Field>
              <Field label={form.objectId ? `${api.t("worker")} (${api.t("object").toLowerCase()})` : api.t("worker")}>
                <select value={form.workerId} onChange={(e) => pickWorker(e.target.value)}>
                  <option value="">{api.t("choose")}</option>
                  {workerOptions.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
                </select>
                {form.objectId && workerOptions.length === 0 && (
                  <div style={{ fontSize: 12, color: S.sub, marginTop: 4 }}>{tr("Nitko još nije upisan na ovom objektu.")}</div>
                )}
              </Field>
              <Field label={api.t("date")}><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></Field>
              <div style={{ marginBottom: 8 }}><DayBadge iso={form.date} /></div>
              <div style={{ display: "flex", gap: 10 }}>
                <div style={{ flex: 1 }}><Field label={api.t("from")}><input type="time" value={form.from} onChange={(e) => setForm({ ...form, from: e.target.value })} /></Field></div>
                <div style={{ flex: 1 }}><Field label={api.t("to")}><input type="time" value={form.to} onChange={(e) => setForm({ ...form, to: e.target.value })} /></Field></div>
              </div>
              <Field label={tr("Napomena (opcionalno)")}><input value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
              <div style={{ background: isDayHoursSuspicious(h) ? S.redSoft : S.greenSoft, borderRadius: 10, padding: "10px 12px", marginBottom: 12, display: "flex", justifyContent: "space-between", fontWeight: 700 }}>
                <span>{api.t("total")}: <span className="num">{fmtH(h)}</span></span>
                {w && <span className="num" style={{ color: S.green }}>{money(round2(h * rateFor(data, w, form.date)), wCur(w))}</span>}
              </div>
              {isDayHoursSuspicious(h) && (
                <div style={{ fontSize: 12.5, color: S.red, fontWeight: 600, margin: "-6px 0 10px" }}>
                  {tr("⚠️ Više od")}{" "}{DAY_HOURS_WARN}{tr("h za jedan dan — provjeri je li ovo greška u unosu.")}
                </div>
              )}
              <Btn onClick={add} style={{ width: "100%" }}>{api.t("addHours")}</Btn>
            </Card>
          ) : (
            <Card>
              <div style={{ display: "flex", gap: 10 }}>
                <div style={{ flex: 1 }}><Field label={tr("Mjesec")}><input type="month" value={monthBulk.month} onChange={(e) => setMonthBulk({ ...monthBulk, month: e.target.value })} /></Field></div>
                <div style={{ flex: 1 }}><Field label={tr("{0} (filter, nije obavezno)", api.t("object"))}><ObjectSelect data={data} api={api} value={monthBulk.objectId} onChange={(v) => setMonthBulk({ ...monthBulk, objectId: v })} /></Field></div>
              </div>
              <div style={{ fontSize: 12, color: S.sub, margin: "-4px 0 10px" }}>{tr("Upiši sate svakome tko je radio i klikni jednom \"Spremi sve\" na kraju — ne treba jedan po jedan.")}</div>
              {monthWorkers.length === 0 ? <Empty text={tr("Nema radnika.")} /> : monthWorkers.map((mw) => (
                <div key={mw.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderBottom: `1px solid ${S.line}` }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 600, fontSize: 14 }}>{mw.name}</div>
                    <div style={{ fontSize: 12, color: S.sub }}>{objName(mw.objectId) || "—"}</div>
                  </div>
                  <input inputMode="decimal" placeholder={tr("sati")} value={monthHours[mw.id] || ""}
                    onChange={(e) => setMonthHours((p) => ({ ...p, [mw.id]: e.target.value }))}
                    style={{ width: 80, padding: "8px 8px", textAlign: "right" }} />
                </div>
              ))}
              <Btn onClick={submitMonthBulk} disabled={monthBusy || monthEntryCount === 0} style={{ width: "100%", marginTop: 12 }}>
                {monthBusy ? tr("Spremam…") : tr("✓ Spremi sve{0}", monthEntryCount ? ` (${monthEntryCount})` : "")}
              </Btn>
              <div style={{ fontSize: 12, color: S.sub, marginTop: 8 }}>{tr("Svaki upis sprema se kao jedan zbirni unos za cijeli mjesec (bez pojedinačnih dana).")}</div>
            </Card>
          )}
        </>
      )}
      {recent.length > 0 && (
        <Card>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{api.t("recent")}</div>
          {recent.map((l) => (
            <div key={l.id} style={{ borderBottom: `1px solid ${S.line}` }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0" }}>
                <div style={{ flex: 1 }}>
                  <div onClick={() => onOpenWorker && onOpenWorker(l.workerId)} style={{ fontWeight: 600, fontSize: 14, cursor: onOpenWorker ? "pointer" : "default", color: onOpenWorker ? S.blue : "inherit" }}>{wName(l.workerId)}</div>
                  <div style={{ fontSize: 12.5, color: S.sub }}>
                    {fmtDate(l.date)} · <LogSpanLabel l={l} />{objName(l.objectId) ? " · " + objName(l.objectId) : ""}{l.note && !l.monthly ? " · " + l.note : ""}
                  </div>
                </div>
                <div className="num" style={{ fontWeight: 700, color: !l.monthly && isDayHoursSuspicious(l.hours) ? S.red : "inherit" }}>
                  {!l.monthly && isDayHoursSuspicious(l.hours) && "⚠️ "}{fmtH(l.hours)}
                </div>
                {(api.admin || l.createdBy === api.uid()) ? (
                  <>
                    <button onClick={() => setEditLog(editLog === l.id ? null : l.id)} style={{ background: "none", border: "none", color: S.blue, fontSize: 15, cursor: "pointer", padding: 4 }}>✎</button>
                    <button onClick={() => api.delLog(l, wName(l.workerId))} style={{ background: "none", border: "none", color: S.red, fontSize: 16, cursor: "pointer", padding: 4 }}>✕</button>
                  </>
                ) : <span title={tr("Upisao kolega")} style={{ fontSize: 13, color: S.sub, padding: 4 }}>👤</span>}
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
function PaymentsTab({ data, api, onOpenWorker }) {
  const [target, setTarget] = useState("radnik"); // radnik | objekt
  const [filterObj, setFilterObj] = useState("");
  const [cur, setCur] = useState("EUR"); // EUR | CZK
  const [form, setForm] = useState({ workerId: "", objectId: "", date: todayISO(), type: "avans", amount: "", note: "", deduct: true });
  const [bulkMode, setBulkMode] = useState(false);
  const [bulkAmounts, setBulkAmounts] = useState({}); // workerId -> string
  const [bulkBusy, setBulkBusy] = useState(false);
  const [pq, setPq] = useState("");
  const wName = (id) => data.workers.find((x) => x.id === id)?.name || "Obrisan radnik";
  const oName = (id) => data.objects.find((o) => o.id === id)?.name || "Obrisan objekt";
  const bulkWorkers = workersAtObject(data, filterObj);
  const bulkEntryCount = Object.values(bulkAmounts).filter((v) => (parseNum(v) || 0) > 0).length;
  const submitBulk = async () => {
    const entries = Object.entries(bulkAmounts)
      .map(([workerId, v]) => ({ workerId, amount: round2(parseNum(v) || 0) }))
      .filter((e) => e.amount > 0);
    if (!entries.length || !form.date) return;
    const deduct = form.type === "avans" || form.type === "racun" ? true : form.type === "bonus" ? false : form.deduct;
    setBulkBusy(true);
    if (await api.addPaymentsBulk(entries, { date: form.date, type: form.type, currency: cur, note: form.note, deduct })) setBulkAmounts({});
    setBulkBusy(false);
  };
  const add = () => {
    const raw = parseNum(form.amount);
    if (!form.date || !raw) return;
    const m = { amount: round2(raw), currency: cur };
    if (target === "radnik") {
      if (!form.workerId) return;
      api.addPayment({ workerId: form.workerId, objectId: "", date: form.date, type: form.type, ...m, note: form.note,
        deduct: form.type === "avans" || form.type === "racun" ? true : form.type === "bonus" ? false : form.deduct }, wName(form.workerId));
    } else {
      if (!form.objectId) return;
      api.addPayment({ workerId: "", objectId: form.objectId, date: form.date, type: form.type === "bonus" || form.type === "racun" ? "ostalo" : form.type,
        ...m, note: form.note, deduct: false }, "🏨 " + oName(form.objectId));
    }
    setForm({ ...form, amount: "", note: "" });
  };
  const recentAll = [...data.payments].sort((a, b) => b.date.localeCompare(a.date));
  const pqTrim = pq.trim().toLowerCase();
  const recent = pqTrim
    ? recentAll.filter((p) => (p.note || "").toLowerCase().includes(pqTrim) || (p.workerId ? wName(p.workerId) : oName(p.objectId)).toLowerCase().includes(pqTrim)).slice(0, 60)
    : recentAll.slice(0, 15);
  const tagStyle = (p) => p.type === "avans" ? { color: S.amber, bg: S.amberSoft }
    : p.type === "bonus" ? { color: S.green, bg: S.greenSoft }
    : p.type === "racun" ? { color: S.blue, bg: S.blueSoft }
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
      {target === "radnik" && (
        <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
          {[[false, "Pojedinačno"], [true, "Grupno — svi odjednom"]].map(([id, label]) => (
            <button key={String(id)} onClick={() => setBulkMode(id)} style={{ flex: 1, padding: "9px 8px", borderRadius: 10, fontWeight: 700, fontSize: 13, cursor: "pointer",
              background: bulkMode === id ? S.green : "#fff", color: bulkMode === id ? "#fff" : S.sub, border: `1px solid ${bulkMode === id ? S.green : S.line}` }}>{label}</button>
          ))}
        </div>
      )}
      {target === "radnik" && data.workers.filter((x) => !x.archived).length === 0 ? <Empty text={tr("Prvo dodaj radnika.")} /> : (
        <Card>
          {target === "radnik" ? (
            <>
              <Field label={tr("Filter po objektu (nije obavezno, samo za lakše pronalaženje)")}>
                <ObjectSelect data={data} api={api} value={filterObj} onChange={(v) => {
                  const stillValid = workersAtObject(data, v).some((x) => x.id === form.workerId);
                  const ob = data.objects.find((o) => o.id === v);
                  setFilterObj(v);
                  setForm({ ...form, workerId: stillValid ? form.workerId : "" });
                  if (v && ob && !stillValid) setCur(countryCur(ob.country));
                }} />
              </Field>
              {!bulkMode && (
                <Field label={filterObj ? `${api.t("worker")} (${api.t("object").toLowerCase()})` : api.t("worker")}>
                  <select value={form.workerId} onChange={(e) => {
                    const wk = data.workers.find((x) => x.id === e.target.value);
                    setForm({ ...form, workerId: e.target.value });
                    if (wk) setCur(wCur(wk));
                  }}>
                    <option value="">{tr("— odaberi —")}</option>
                    {workersAtObject(data, filterObj).map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
                  </select>
                  {filterObj && workersAtObject(data, filterObj).length === 0 && (
                    <div style={{ fontSize: 12, color: S.sub, marginTop: 4 }}>{tr("Nitko još nije upisan na ovom objektu.")}</div>
                  )}
                </Field>
              )}
              <Field label={api.t("type")}>
                <select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
                  <option value="avans">{tr("Avans (odbija se od plaće)")}</option>
                  <option value="bonus">{tr("Bonus (dodaje se na plaću)")}</option>
                  <option value="racun">{tr("Isplaćeno na račun (banka) — ostatak ide na kovertu")}</option>
                  <option value="gorivo">{tr("Gorivo")}</option>
                  <option value="ostalo">{tr("Ostali trošak")}</option>
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
                <select value={form.type === "bonus" || form.type === "racun" ? "ostalo" : form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
                  <option value="avans">{tr("Avans")}</option>
                  <option value="gorivo">{tr("Gorivo")}</option>
                  <option value="ostalo">{tr("Ostali trošak")}</option>
                </select>
              </Field>
              <div style={{ fontSize: 12.5, color: S.sub, margin: "-4px 0 10px" }}>
                {tr("Trošak se veže na objekt (ne odbija se nijednom radniku) i admin ga vidi u dobiti objekta.")}
              </div>
            </>
          )}
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}><Field label={api.t("date")}><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></Field></div>
            {!(target === "radnik" && bulkMode) && (
              <div style={{ flex: 1 }}>
                <Field label={cur === "CZK" ? tr("Iznos (Kč)") : api.t("amount")}>
                  <input inputMode="decimal" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} placeholder="0.00" />
                </Field>
              </div>
            )}
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10 }}>
            {[["EUR", "€ Euro"], ["CZK", "Kč Kruna"]].map(([id, label]) => (
              <button key={id} onClick={() => setCur(id)} style={{ padding: "7px 14px", borderRadius: 999, fontWeight: 700, fontSize: 13, cursor: "pointer",
                background: cur === id ? S.green : "#fff", color: cur === id ? "#fff" : S.sub, border: `1px solid ${cur === id ? S.green : S.line}` }}>{label}</button>
            ))}
          </div>
          <Field label={form.type === "bonus" ? tr("Za što je bonus") : tr("Napomena")}>
            <input value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })}
              placeholder={form.type === "bonus" ? tr("npr. dobro odrađena sezona…") : form.type === "racun" ? tr("npr. uplata na IBAN") : tr("npr. gorivo za put")} />
          </Field>
          {target === "radnik" && form.type !== "avans" && form.type !== "bonus" && form.type !== "racun" && (
            <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, fontSize: 14 }}>
              <input type="checkbox" checked={form.deduct} onChange={(e) => setForm({ ...form, deduct: e.target.checked })} style={{ width: 18, height: 18 }} />
              {tr("Odbij radniku od plaće (inače je trošak firme)")}
            </label>
          )}
          {target === "radnik" && bulkMode ? (
            <>
              <div style={{ fontSize: 12, color: S.sub, margin: "-4px 0 10px" }}>{tr("Upiši iznos svakome tko treba i klikni jednom \"Spremi sve\" na kraju.")}</div>
              {bulkWorkers.length === 0 ? <Empty text={tr("Nema radnika.")} /> : bulkWorkers.map((bw) => (
                <div key={bw.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderBottom: `1px solid ${S.line}` }}>
                  <div style={{ flex: 1, fontWeight: 600, fontSize: 14 }}>{bw.name}</div>
                  <input inputMode="decimal" placeholder="0.00" value={bulkAmounts[bw.id] || ""}
                    onChange={(e) => setBulkAmounts((p) => ({ ...p, [bw.id]: e.target.value }))}
                    style={{ width: 90, padding: "8px 8px", textAlign: "right" }} />
                </div>
              ))}
              <Btn onClick={submitBulk} disabled={bulkBusy || bulkEntryCount === 0} style={{ width: "100%", marginTop: 12 }}>
                {bulkBusy ? tr("Spremam…") : tr("✓ Spremi sve{0}", bulkEntryCount ? ` (${bulkEntryCount})` : "")}
              </Btn>
            </>
          ) : (
            <Btn onClick={add} style={{ width: "100%" }}>{api.t("save")}</Btn>
          )}
        </Card>
      )}
      {recentAll.length > 15 && (
        <input value={pq} onChange={(e) => setPq(e.target.value)} placeholder={tr("🔍 Traži stariji upis po imenu ili napomeni (za brisanje)…")} style={{ marginBottom: 10 }} />
      )}
      {recent.length > 0 && (
        <Card>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>{pqTrim ? tr("Rezultati ({0})", recent.length) : api.t("recent")}</div>
          {recent.map((p) => {
            const t = tagStyle(p);
            return (
              <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: `1px solid ${S.line}` }}>
                <Tag color={t.color} bg={t.bg}>{tr(TYPE_LABEL[p.type] || p.type)}</Tag>
                <div style={{ flex: 1 }}>
                  <div onClick={() => p.workerId && onOpenWorker && onOpenWorker(p.workerId)}
                    style={{ fontWeight: 600, fontSize: 14, cursor: p.workerId && onOpenWorker ? "pointer" : "default", color: p.workerId && onOpenWorker ? S.blue : "inherit" }}>
                    {p.workerId ? wName(p.workerId) : "🏨 " + oName(p.objectId)}
                  </div>
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

// način isplate izveden iz već upisanih isplata na račun (bez posebnog polja) — koristi se samo za prikaz
const paymentMethodLabel = (r) => {
  const isCzk = wCur(r.w) === "CZK";
  const bank = isCzk ? r.czk.bank : r.bank;
  const net = isCzk ? r.czk.net : r.net;
  if (!bank) return "💶 Gotovina";
  if (net <= 0.01) return "🏦 Na račun";
  return "🏦+💶 Dio na račun";
};

/* ================================================================== */
/*  OBRAČUN: mjesec/godina, isplaćeno, PDF, naplata, grafovi           */
/* ================================================================== */
function calcRows(data, filterFn, objFilters) {
  const hasFilter = objFilters && objFilters.size > 0;
  const matchObj = (oid) => !hasFilter || objFilters.has(oid);
  return data.workers.map((w) => {
    const logs = data.logs.filter((l) => l.workerId === w.id && filterFn(l.date) && matchObj(l.objectId));
    const pays = data.payments.filter((p) => p.workerId === w.id && filterFn(p.date));
    const hours = round2(logs.reduce((s, l) => s + l.hours, 0));
    const grossAll = round2(logs.reduce((s, l) => s + l.hours * rateFor(data, w, l.date), 0));
    const gross = wCur(w) === "CZK" ? 0 : grossAll;
    const grossKc = wCur(w) === "CZK" ? grossAll : 0;
    const extraAll = round2(logs.reduce((s, l) => s + extraPayForLog(data, w, l), 0));
    const extraPay = wCur(w) === "CZK" ? 0 : extraAll;
    const extraPayKc = wCur(w) === "CZK" ? extraAll : 0;
    const rateSet = [...new Set(logs.map((l) => rateFor(data, w, l.date)))];
    const by = (c) => {
      const mine = pays.filter((p) => (p.currency || "EUR") === c);
      return {
        bonuses: round2(mine.filter((p) => p.type === "bonus").reduce((s, p) => s + p.amount, 0)),
        advances: round2(mine.filter((p) => p.type === "avans").reduce((s, p) => s + p.amount, 0)),
        bank: round2(mine.filter((p) => p.type === "racun").reduce((s, p) => s + p.amount, 0)),
        deductions: round2(mine.filter((p) => p.type !== "avans" && p.type !== "bonus" && p.type !== "racun" && p.deduct).reduce((s, p) => s + p.amount, 0)),
        firmCosts: round2(mine.filter((p) => p.type !== "avans" && p.type !== "bonus" && p.type !== "racun" && !p.deduct).reduce((s, p) => s + p.amount, 0)),
      };
    };
    const e = by("EUR"), k = by("CZK");
    return { w, logs, pays, hours, gross, grossKc, rateSet, extraPay, extraPayKc,
      bonuses: e.bonuses, advances: e.advances, deductions: e.deductions, firmCosts: e.firmCosts, bank: e.bank,
      net: round2(gross + extraPay + e.bonuses - e.advances - e.deductions - e.bank),
      czk: { ...k, gross: grossKc, net: round2(grossKc + extraPayKc + k.bonuses - k.advances - k.deductions - k.bank) } };
  }).filter((r) => (hasFilter ? r.hours > 0 : (r.hours > 0 || r.pays.length > 0)));
}

function ReportTab({ data, api, admin, onOpenWorker }) {
  const [view, setView] = useState("month"); // month | year
  const [month, setMonth] = useState(workMonthDefault());
  const [open, setOpen] = useState(null);
  const [uplata, setUplata] = useState({});
  const [objFilters, setObjFilters] = useState(() => new Set()); // prazan skup = svi objekti
  const [objPickerOpen, setObjPickerOpen] = useState(false);
  const [showPayoutCal, setShowPayoutCal] = useState(false);
  const [showCombinedEur, setShowCombinedEur] = useState(false);
  const [section, setSection] = useState("radnici"); // radnici | objekti | analiza
  const [exportOpen, setExportOpen] = useState(false);
  const [payStatusFilter, setPayStatusFilter] = useState("all"); // all | paid | pending | unpaid
  const [signatures, setSignatures] = useState({}); // workerId -> dataURL, samo za trenutnu sesiju (ne sprema se u bazu)
  const [sigTarget, setSigTarget] = useState(null); // { id, name } radnika koji trenutno potpisuje
  const [selected, setSelected] = useState(() => new Set());
  const [bulkBusy, setBulkBusy] = useState(false);
  const [mzdyBusy, setMzdyBusy] = useState(false);
  const [mzdyErr, setMzdyErr] = useState("");
  const [mzdyReview, setMzdyReview] = useState(null); // { month, rows: [...] }
  const mzdyFileRef = useRef(null);
  const [y, m] = month.split("-").map(Number);
  const prev = () => setMonth(m === 1 ? `${y - 1}-12` : `${y}-${String(m - 1).padStart(2, "0")}`);
  const next = () => setMonth(m === 12 ? `${y + 1}-01` : `${y}-${String(m + 1).padStart(2, "0")}`);
  const objName = (id) => data.objects.find((o) => o.id === id)?.name || "";
  const objFilterName = objFilters.size === 0 ? "" : [...objFilters].map(objName).join(", ");
  const objFilterLabel = objFilters.size === 0 ? api.t("allObjects")
    : objFilters.size <= 3 ? [...objFilters].map(objName).join(", ")
    : `${objFilters.size} objekata odabrano`;
  const toggleObjFilter = (id) => setObjFilters((prev) => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });

  const inMonth = (d) => monthKey(d) === month;
  useEffect(() => { setSelected(new Set()); }, [month, view, objFilters]);
  const inYear = (d) => (d || "").slice(0, 4) === String(y);
  const rows = calcRows(data, view === "month" ? inMonth : inYear, objFilters).sort((a, b) => a.w.name.localeCompare(b.w.name, "hr"));
  const payStatusOf = (r) => {
    const p = view === "month" ? paidFor(data, r.w.id, month) : null;
    return p ? (p.approved ? "paid" : "pending") : "unpaid";
  };
  const displayRows = view === "month" && payStatusFilter !== "all" ? rows.filter((r) => payStatusOf(r) === payStatusFilter) : rows;
  const exportTargetRows = selected.size > 0 ? displayRows.filter((r) => selected.has(r.w.id)) : displayRows;
  const PAY_STATUS_LABEL = { all: "", paid: "✓ Isplaćeno", pending: "⏳ Predloženo", unpaid: "Bez isplate" };
  const exportScopeLabel = selected.size > 0 ? `${selected.size} odabrano`
    : (view === "month" ? PAY_STATUS_LABEL[payStatusFilter] : "");
  const exportScopeText = [
    selected.size > 0 ? tr("Za {0} označenih radnika", selected.size)
      : payStatusFilter !== "all" && view === "month" ? tr("Za filter: {0}", tr(PAY_STATUS_LABEL[payStatusFilter]))
      : tr("Za sve prikazane radnike"),
    objFilterName,
  ].filter(Boolean).join(" · ");
  const unpaidRows = view === "month" ? displayRows.filter((r) => !paidFor(data, r.w.id, month)) : [];
  const pendingRows = view === "month" ? displayRows.filter((r) => { const p = paidFor(data, r.w.id, month); return p && !p.approved; }) : [];
  const allSelected = unpaidRows.length > 0 && unpaidRows.every((r) => selected.has(r.w.id));
  const toggleSelectAll = () => setSelected((prev) => {
    const n = new Set(prev);
    unpaidRows.forEach((r) => (allSelected ? n.delete(r.w.id) : n.add(r.w.id)));
    return n;
  });
  const allPendingSelected = pendingRows.length > 0 && pendingRows.every((r) => selected.has(r.w.id));
  const toggleSelectAllPending = () => setSelected((prev) => {
    const n = new Set(prev);
    pendingRows.forEach((r) => (allPendingSelected ? n.delete(r.w.id) : n.add(r.w.id)));
    return n;
  });
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
  const payAndPrintMany = async (list) => {
    if (!list.length) return;
    printEnvelopes(list);
    await payMany(list);
  };
  const approveMany = async (list) => {
    setBulkBusy(true);
    for (const r of list) {
      const p = paidFor(data, r.w.id, month);
      if (p && !p.approved) await api.approvePayout(p, r.w.name);
    }
    setSelected(new Set());
    setBulkBusy(false);
  };

  const czkRate = Number(data.settings.czk_rate) || 25;
  const handleMzdyFile = async (e) => {
    const files = [...(e.target.files || [])];
    e.target.value = "";
    if (!files.length) return;
    setMzdyErr(""); setMzdyBusy(true);
    try {
      // "oprava" se obrađuje zadnja → zamjenjuje stare pásky istog radnika
      files.sort((a, b) => (/oprav/i.test(a.name) ? 1 : 0) - (/oprav/i.test(b.name) ? 1 : 0));
      const people = new Map(); // normName → páska
      let naklady = [], prikazy = [];
      const docs = [];
      for (const f of files) {
        const lines = await extractPdfLines(f);
        const type = mzdyDocType(lines);
        if (type === "naklady") { naklady = parseMzdyNaklady(lines); docs.push(`${f.name}: mzdové náklady`); }
        else if (type === "prikazy") { prikazy = parseMzdyPrikazy(lines); docs.push(`${f.name}: příkazy k úhradě`); }
        else {
          const list = parseMzdyPasky(lines);
          list.forEach((p) => people.set(normName(p.name), { ...p, corrected: people.has(normName(p.name)) }));
          docs.push(`${f.name}: ${list.length} pásky${/oprav/i.test(f.name) ? " (oprava)" : ""}`);
        }
      }
      if (!people.size) { setMzdyErr(tr("Nisam našao nijednu výplatní pásku. Učitaj PDF s páskama (i opravu ako postoji).")); setMzdyBusy(false); return; }
      const monthCounts = {};
      people.forEach((p) => { monthCounts[p.month] = (monthCounts[p.month] || 0) + 1; });
      const mo = Object.entries(monthCounts).sort((a, b) => b[1] - a[1])[0][0];
      const rows = [...people.values()].map((p) => {
        const w = matchWorkerFuzzy(data.workers, p.name);
        const cur = w ? wCur(w) : "CZK";
        const conv = (kc) => String(cur === "CZK" ? kc : round2(kc / czkRate));
        const soc = socZdrFor(p);
        const nk = naklady.find((n) => plainTokens(p.name)[0] === plainTokens(n.surname)[0] && plainTokens(p.name).slice(1).join(" ").startsWith(plainTokens(n.first).join(" ").slice(0, 5)));
        const exBank = w ? round2(data.payments.filter((x) => x.workerId === w.id && x.type === "racun" && monthKey(x.date) === mo && (x.currency || "EUR") === cur && !/Mzdy PDF/.test(x.note || "")).reduce((s, x) => s + x.amount, 0)) : 0;
        const bankAmt = Number(conv(p.dobirka));
        const exNote = w ? (data.payrollNotes || []).filter((n) => n.workerId === w.id && n.month === mo).reduce((s, n) => s + n.amount, 0) : 0;
        const locked = w ? !!paidFor(data, w.id, mo)?.approved : false;
        return {
          rawName: p.name, kat: p.kat, tarif: p.tarif, corrected: p.corrected, workerId: w?.id || "", currency: cur,
          socZdr: conv(soc), bank: conv(p.dobirka), dobirkaKc: p.dobirka,
          include: !!w && !locked, addBank: !!w && p.dobirka > 0 && !(exBank > 0),
          exBank, exNote, locked,
          warn: [
            nk && Math.round(nk.dobirka) !== Math.round(p.dobirka) ? `Mzdové náklady kažu ${czk(nk.dobirka)} na račun` : "",
            exBank > 0 ? `već ručno upisano na račun ${money(exBank, cur)}${Math.round(exBank) === Math.round(bankAmt) ? " (isto)" : ""}` : "",
            locked ? "mjesec je odobren i zaključan" : "",
          ].filter(Boolean),
        };
      }).sort((a, b) => (a.workerId ? 0 : 1) - (b.workerId ? 0 : 1) || a.rawName.localeCompare(b.rawName, "hr"));
      setMzdyReview({ month: mo, rows, prikazy, docs, addReminders: prikazy.length > 0 });
    } catch (err) {
      setMzdyErr(tr("Ne mogu pročitati datoteku: ") + (err.message || err));
    }
    setMzdyBusy(false);
  };
  const saveMzdyReview = async () => {
    const mo = mzdyReview.month;
    const [yy, mm] = mo.split("-").map(Number);
    const lastDay = `${mo}-${String(new Date(yy, mm, 0).getDate()).padStart(2, "0")}`;
    const nextMonth20 = (() => { const d = new Date(yy, mm, 20); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-20`; })();
    const rows = mzdyReview.rows.filter((r) => r.include && r.workerId && !r.locked);
    if (!rows.length && !(mzdyReview.addReminders && mzdyReview.prikazy.length)) { setMzdyErr(tr("Ništa nije odabrano za spremanje.")); return; }
    setMzdyBusy(true); setMzdyErr("");
    const uid = api.uid();
    const tag = `Dobírka (Mzdy PDF ${String(mm).padStart(2, "0")}/${yy})`;
    const ok = await api.run(async () => {
      for (const r of rows) {
        const soc = round2(parseNum(r.socZdr) || 0);
        // soc.-zdr.: stara napomena za taj mjesec se zamjenjuje novom
        const old = (data.payrollNotes || []).filter((n) => n.workerId === r.workerId && n.month === mo);
        if (old.length) {
          const { error } = await supabase.from("payroll_notes").update({ deleted_at: new Date().toISOString() }).in("id", old.map((n) => n.id));
          if (error) throw error;
        }
        if (soc > 0) {
          const { error } = await supabase.from("payroll_notes").insert({ worker_id: r.workerId, month: mo, amount: soc, currency: r.currency,
            source_czk: r.currency === "CZK" ? soc : round2(soc * czkRate), note: MZDY_NOTE_LABEL, created_by: uid });
          if (error) throw error;
        }
        // na račun (Dobírka): prethodni uvoz iz PDF-a se zamjenjuje, ručni upisi ostaju
        const bank = round2(parseNum(r.bank) || 0);
        if (r.addBank && bank > 0) {
          const prev = data.payments.filter((x) => x.workerId === r.workerId && x.type === "racun" && monthKey(x.date) === mo && /Mzdy PDF/.test(x.note || ""));
          if (prev.length) {
            const { error } = await supabase.from("payments").update({ deleted_at: new Date().toISOString() }).in("id", prev.map((x) => x.id));
            if (error) throw error;
          }
          const { error } = await supabase.from("payments").insert({ worker_id: r.workerId, object_id: null, pay_date: lastDay, type: "racun",
            amount: bank, note: tag, deduct: true, currency: r.currency, created_by: uid });
          if (error) throw error;
        }
      }
      if (mzdyReview.addReminders) {
        for (const p of mzdyReview.prikazy) {
          const text = `Platiti ${p.to || ""} ${czk(p.amount)} – ${p.label} ${String(mm).padStart(2, "0")}/${yy}`.replace(/\s+/g, " ").trim();
          if ((data.reminders || []).some((x) => x.text === text)) continue;
          const { error } = await supabase.from("reminders").insert({ text, due_date: nextMonth20, created_by: uid });
          if (error) throw error;
        }
      }
    }, `Uvezao Mzdy PDF za ${mo}: ${rows.length} radnika (soc.-zdr. i na račun)`);
    setMzdyBusy(false);
    if (ok) setMzdyReview(null);
  };

  const periodLogs = data.logs.filter((l) => (view === "month" ? inMonth : inYear)(l.date) && (objFilters.size === 0 || objFilters.has(l.objectId)));
  const periodBrigade = (data.brigadeCosts || []).filter((bg) => (view === "month" ? inMonth : inYear)(bg.date) && (objFilters.size === 0 || objFilters.has(bg.objectId)));
  const periodLabel = view === "month" ? `${MONTHS[m - 1]} ${y}.` : `${y}. godina`;
  const periodLabelUi = view === "month" ? `${mName(m - 1)} ${y}.` : tr("{0}. godina", y);

  const totals = rows.reduce((t, r) => ({
    hours: round2(t.hours + r.hours), gross: round2(t.gross + r.gross),
    net: round2(t.net + r.net), firm: round2(t.firm + r.firmCosts), bonus: round2(t.bonus + r.bonuses),
    netKc: round2(t.netKc + r.czk.net), firmKc: round2(t.firmKc + r.czk.firmCosts),
    grossKc: round2(t.grossKc + r.grossKc), bonusKc: round2(t.bonusKc + r.czk.bonuses),
  }), { hours: 0, gross: 0, net: 0, firm: 0, bonus: 0, netKc: 0, firmKc: 0, grossKc: 0, bonusKc: 0 });

  const objPays = data.payments.filter((p) => !p.workerId && p.objectId && (view === "month" ? inMonth : inYear)(p.date));
  const brigadeCostsTotal = round2(periodBrigade.filter((bg) => (bg.currency || "CZK") === "EUR").reduce((s, bg) => s + bg.amount, 0));
  const brigadeCostsTotalKc = round2(periodBrigade.filter((bg) => bg.currency === "CZK").reduce((s, bg) => s + bg.amount, 0));
  const objCostsTotal = round2(objPays.filter((p) => (p.currency || "EUR") === "EUR").reduce((s, p) => s + p.amount, 0) + brigadeCostsTotal);
  const objCostsTotalKc = round2(objPays.filter((p) => p.currency === "CZK").reduce((s, p) => s + p.amount, 0) + brigadeCostsTotalKc);
  const byObject = new Map();
  periodLogs.forEach((l) => {
    const key = l.objectId || "__none__";
    const wk = data.workers.find((x) => x.id === l.workerId);
    const ob = data.objects.find((o) => o.id === l.objectId);
    const cur = byObject.get(key) || { hours: 0, gross: 0, grossKc: 0, revenue: 0, revenueKc: 0, costs: 0, costsKc: 0 };
    cur.hours = round2(cur.hours + l.hours);
    const gv = l.hours * rateFor(data, wk, l.date) + extraPayForLog(data, wk, l);
    if (wCur(wk) === "CZK") cur.grossKc = round2(cur.grossKc + gv); else cur.gross = round2(cur.gross + gv);
    const { rate: brate, cur: bcur } = billRateForLog(data, ob, wk, l.date);
    const rv = l.hours * brate;
    if (bcur === "CZK") cur.revenueKc = round2(cur.revenueKc + rv); else cur.revenue = round2(cur.revenue + rv);
    byObject.set(key, cur);
  });
  periodBrigade.forEach((bg) => {
    const ob = data.objects.find((o) => o.id === bg.objectId);
    const cur = byObject.get(bg.objectId) || { hours: 0, gross: 0, grossKc: 0, revenue: 0, revenueKc: 0, costs: 0, costsKc: 0 };
    cur.hours = round2(cur.hours + bg.hours);
    // posebna cijena te brigade (ako je upisana) nadjačava zadanu cijenu objekta
    const brate = bg.billRate > 0 ? bg.billRate : (ob?.billRate || 0);
    const bcur2 = bg.billRate > 0 ? bg.billCur : (ob?.billCur || "EUR");
    const rv = round2(bg.hours * brate);
    if (bcur2 === "CZK") cur.revenueKc = round2(cur.revenueKc + rv); else cur.revenue = round2(cur.revenue + rv);
    if (bg.currency === "CZK") cur.costsKc = round2((cur.costsKc || 0) + bg.amount); else cur.costs = round2((cur.costs || 0) + bg.amount);
    byObject.set(bg.objectId, cur);
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
        const { rate: brate, cur: bcur } = billRateForLog(data, ob, wk, l.date);
        const rv = l.hours * brate;
        if (bcur === "CZK") revK += rv; else revE += rv;
        const cv = l.hours * rateFor(data, wk, l.date) + extraPayForLog(data, wk, l);
        if (wCur(wk) === "CZK") costK2 += cv; else costE2 += cv;
      });
      const extraOf = (c) => data.payments.filter((p) => monthKey(p.date) === mk2 && (p.currency || "EUR") === c && (!p.workerId ? !!p.objectId : (p.type === "bonus" || (!p.deduct && p.type !== "avans")))).reduce((s, p) => s + p.amount, 0);
      profit = round2(revE - costE2 - extraOf("EUR"));
      profitKc = round2(revK - costK2 - extraOf("CZK"));
    }
    last6.push({ label: mShort(d.getMonth()), hours: hrs, profit, profitKc });
  }

  /* razrada obračuna jednog radnika — ista za ekran i PDF (oznake su hrvatske, na ekranu idu kroz tr) */
  const noteFor = (r) => (view === "month" ? (data.payrollNotes || []).find((n) => n.workerId === r.w.id && n.month === month) : null);
  const amtText = (e, k, gc) => [e ? eur(e) : "", k ? czk(k) : ""].filter(Boolean).join(" + ") || money(0, gc);
  const breakdownOf = (r) => {
    const note = noteFor(r);
    const gc = wCur(r.w);
    const items = [
      { label: "Zarada", hint: fmtH(r.hours), e: r.gross, k: r.grossKc, sign: "", always: true },
      { label: "Dodatak vikend / praznik", e: r.extraPay, k: r.extraPayKc, sign: "+" },
      { label: "Bonus", e: r.bonuses, k: r.czk.bonuses, sign: "+" },
      { label: "Avans", e: r.advances, k: r.czk.advances, sign: "−" },
      { label: "Odbici", e: r.deductions, k: r.czk.deductions, sign: "−" },
      { label: "Na račun (banka)", e: r.bank, k: r.czk.bank, sign: "−" },
      note ? { label: note.note || MZDY_NOTE_LABEL, e: note.currency === "EUR" ? note.amount : 0, k: note.currency === "CZK" ? note.amount : 0, sign: "−" } : null,
    ].filter((x) => x && (x.always || x.e || x.k));
    const netE = round2(r.net - (note && note.currency === "EUR" ? note.amount : 0));
    const netK = round2(r.czk.net - (note && note.currency === "CZK" ? note.amount : 0));
    return { items, netE, netK, note, gc, netText: amtText(netE, netK, gc) };
  };

  /* PDF obračun za radnika */
  const payslipBody = (r, st) => {
    const gc = wCur(r.w);
    const bd = breakdownOf(r);
    const docNo = `${view === "month" ? `${y}-${String(m).padStart(2, "0")}` : y}/${String(r.w.id).replace(/-/g, "").slice(0, 6).toUpperCase()}`;
    const logs = [...r.logs].sort((a, b) => a.date.localeCompare(b.date));
    const logsHtml = logs.map((l) =>
      `<tr><td>${fmtDate(l.date)}</td><td>${logSpan(l)}</td><td>${objName(l.objectId) || ""}</td><td class="right">${fmtH(l.hours)}</td><td class="right">${money(round2(l.hours * rateFor(data, r.w, l.date) + extraPayForLog(data, r.w, l)), gc)}</td></tr>`).join("");
    const logsTotal = round2((gc === "CZK" ? r.grossKc + r.extraPayKc : r.gross + r.extraPay));
    const itemsHtml = bd.items.map((it) =>
      `<tr><td>${it.label}${it.hint ? ` (${it.hint})` : ""}</td><td class="right">${it.sign}${amtText(it.e, it.k, gc)}</td></tr>`).join("");
    const paysHtml = r.pays.length ? `
      <h3>Isplate i odbici u razdoblju</h3>
      <table><tr><th style="width:90px">Datum</th><th>Vrsta</th><th>Napomena</th><th class="right">Iznos</th></tr>
      ${[...r.pays].sort((a, b) => a.date.localeCompare(b.date)).map((p) =>
        `<tr><td>${fmtDate(p.date)}</td><td>${TYPE_LABEL[p.type] || p.type}</td><td>${p.note || ""}</td><td class="right">${p.type === "bonus" ? "+" : p.deduct ? "−" : ""}${money(p.amount, p.currency)}</td></tr>`).join("")}</table>` : "";
    return `
      ${firmHead(st)}
      <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:16px">
        <div><h1>Pregled sati i isplate</h1><h2>${r.w.name}${r.w.position ? " · " + r.w.position : ""} · ${periodLabel}</h2></div>
        <div class="muted" style="text-align:right;white-space:nowrap">Broj: <b>${docNo}</b><br>Datum izdavanja: ${fmtDate(todayISO())}</div>
      </div>
      <h3>Evidencija sati</h3>
      <table><tr><th style="width:90px">Datum</th><th>Vrijeme</th><th>Objekt</th><th class="right">Sati</th><th class="right">Iznos</th></tr>
      ${logsHtml || `<tr><td colspan="5" class="muted">Nema upisanih sati.</td></tr>`}
      <tr class="tot"><td colspan="3">Ukupno</td><td class="right">${fmtH(r.hours)}</td><td class="right">${money(logsTotal, gc)}</td></tr></table>
      ${paysHtml}
      <h3>Obračun</h3>
      <table style="max-width:420px;margin-left:auto">${itemsHtml}
      <tr class="tot"><td>Za isplatu (gotovina)</td><td class="right">${bd.netText}</td></tr></table>
      <div style="margin-top:40px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:24px;font-size:12px;color:#5C6862">
        <div><div style="border-bottom:1px solid #1C2521;height:60px"></div>Mjesto i datum</div>
        <div><div style="border-bottom:1px solid #1C2521;height:60px;display:flex;align-items:flex-end">${signatures[r.w.id] ? `<img src="${signatures[r.w.id]}" style="height:56px" />` : ""}</div>Potpis radnika</div>
        <div><div style="border-bottom:1px solid #1C2521;height:60px"></div>Potpis poslodavca</div>
      </div>
      <div class="muted" style="margin-top:28px;font-size:10.5px">Interni pregled odrađenih sati i gotovinske isplate. Ne zamjenjuje službenu výplatní pásku.</div>`;
  };
  const printPayslip = (r) => printDoc(`Obračun ${r.w.name}`, payslipBody(r, api.settings));
  // dijeli tekstualni sažetak (ne i sam PDF — pravo dijeljenje PDF datoteke traži dodatnu biblioteku)
  const shareSummary = async (r) => {
    const text = `Obračun — ${r.w.name}\n${periodLabel}\nSati: ${fmtH(r.hours)}\nZa isplatu: ${breakdownOf(r).netText}`;
    if (navigator.share) { try { await navigator.share({ title: `Obračun ${r.w.name}`, text }); } catch {} }
    else if (navigator.clipboard) { try { await navigator.clipboard.writeText(text); alert(tr("Dijeljenje nije podržano u ovom pregledniku — sažetak je kopiran u međuspremnik.")); } catch { alert(text); } }
    else alert(text);
  };
  const printAllPayslips = () => {
    const st = api.settings;
    const body = exportTargetRows.map((r) => `<div style="page-break-after:always">${payslipBody(r, st)}</div>`).join("");
    printDoc(`Svi obračuni ${periodLabel}`, body);
  };

  /* PDF specifikacija za hotel */
  const printInvoiceSpec = (key, v) => {
    const st = api.settings;
    const ob = data.objects.find((o) => o.id === key);
    const ls = periodLogs.filter((l) => l.objectId === key).sort((a, b) => a.date.localeCompare(b.date));
    const brig = periodBrigade.filter((bg) => bg.objectId === key).sort((a, b) => a.date.localeCompare(b.date));
    const rowsHtml = ls.map((l) => {
      const wk = data.workers.find((x) => x.id === l.workerId);
      return `<tr><td>${fmtDate(l.date)}</td><td>${wk?.name || ""}${wk?.position ? " · " + wk.position : ""}</td><td>${logSpan(l)}</td><td class="right">${fmtH(l.hours)}</td></tr>`;
    }).join("") + brig.map((bg) =>
      `<tr><td>${fmtDate(bg.date)}</td><td>Brigada (bez imena)</td><td>—</td><td class="right">${fmtH(bg.hours)}</td></tr>`).join("");
    const rv = objRevenue(data, ob, ls, brig);
    const priceRowsHtml = rv.byPosition.map((row) =>
      `<tr><td>${row.position || "Naplata"}: ${fmtH(row.hours)} × ${money(row.rate, row.cur)}/h</td><td class="right">${money(row.amount, row.cur)}</td></tr>`).join("");
    const totalRowsHtml = [
      rv.revenue ? `<tr class="tot"><td>ZA NAPLATU (€)</td><td class="right">${eur(rv.revenue)}</td></tr>` : "",
      rv.revenueKc ? `<tr class="tot"><td>ZA NAPLATU (Kč)</td><td class="right">${czk(rv.revenueKc)}</td></tr>` : "",
    ].join("");
    printDoc(`Specifikacija ${ob?.name}`, `
      ${firmHead(st)}
      <h1>Specifikacija sati — ${ob?.name}</h1><h2>${periodLabel}</h2>
      <table><tr><th>Datum</th><th>Radnik</th><th>Vrijeme</th><th class="right">Sati</th></tr>${rowsHtml}
      <tr class="tot"><td colspan="3">UKUPNO</td><td class="right">${fmtH(v.hours)}</td></tr></table>
      <table><tr><th></th><th class="right">Iznos</th></tr>
      ${priceRowsHtml}
      ${totalRowsHtml}</table>
      ${st.iban ? `<div class="muted">Uplata na IBAN: <b>${st.iban}</b></div>` : ""}`);
  };

  /* ---------- Koverte za isplatu (jedna po radniku, za ispis) ---------- */
  const printEnvelopes = (list) => {
    const pages = (list || rows).map((r) => {
      const objs = [...new Set(r.logs.map((l) => objName(l.objectId)).filter(Boolean))];
      const objLine = objs.length ? objs.join(", ") : (objName(r.w.objectId) || "");
      const note = view === "month" ? (data.payrollNotes || []).find((n) => n.workerId === r.w.id && n.month === month) : null;
      const netE = round2(r.net - (note && note.currency === "EUR" ? note.amount : 0));
      const netK = round2(r.czk.net - (note && note.currency === "CZK" ? note.amount : 0));
      if (netE === 0 && netK === 0) return "";
      const amounts = [
        netE !== 0 || netK === 0 ? eur(netE) : "",
        netK !== 0 ? czk(netK) : "",
      ].filter(Boolean);
      const gc = wCur(r.w);
      const noteLabel = note?.note || MZDY_NOTE_LABEL;
      const breakdown = [
        gc === "CZK" ? (r.grossKc ? "zarada " + czk(r.grossKc) : "") : (r.gross ? "zarada " + eur(r.gross) : ""),
        r.extraPay ? "vikend/praznik +" + eur(r.extraPay) : "", r.extraPayKc ? "vikend/praznik +" + czk(r.extraPayKc) : "",
        r.bonuses ? "bonus +" + eur(r.bonuses) : "", r.czk.bonuses ? "bonus +" + czk(r.czk.bonuses) : "",
        r.advances ? "avans −" + eur(r.advances) : "", r.czk.advances ? "avans −" + czk(r.czk.advances) : "",
        r.deductions ? "odbici −" + eur(r.deductions) : "", r.czk.deductions ? "odbici −" + czk(r.czk.deductions) : "",
        r.bank ? "na račun −" + eur(r.bank) : "", r.czk.bank ? "na račun −" + czk(r.czk.bank) : "",
        note ? `${noteLabel} −${money(note.amount, note.currency)}` : "",
      ].filter(Boolean).join(" · ");
      const envPeriodLabel = view === "month" ? `${m}. ${y}.` : `${y}.`;
      return `<div class="env">
        <div class="env-top">${envPeriodLabel}${objLine ? "<br>" + objLine : ""}</div>
        <div class="env-mid">
          <div class="env-name">${r.w.name}</div>
          <div class="env-hours">${fmtH(r.hours)}</div>
        </div>
        ${breakdown && breakdown !== ("zarada " + amounts[0]) ? `<div class="env-break">${breakdown}</div>` : ""}
        <div class="env-total">${amounts.join("<br>")}</div>
      </div>`;
    }).join("");
    printDoc("Koverte " + periodLabel + (objFilterName ? " — " + objFilterName : ""), `<style>
      @page{size:220mm 110mm;margin:0}
      body{padding:0!important;margin:0!important;max-width:none!important}
      .env{width:220mm;box-sizing:border-box;padding:10mm 18mm 0;page-break-after:always;page-break-inside:avoid;text-align:center}
      .env-top{font-size:12pt;color:#444;line-height:1.3}
      .env-mid{margin:2mm 0 1.5mm}
      .env-name{font-size:19pt;font-weight:bold;letter-spacing:.5px;margin-bottom:1mm;text-transform:uppercase}
      .env-hours{font-size:11.5pt}
      .env-break{font-size:9pt;color:#444;line-height:1.25;margin-bottom:1mm}
      .env-total{border-top:2px solid #000;width:70%;margin:4mm auto 0;padding-top:2mm;font-size:17pt;font-weight:bold;line-height:1.2}
    </style>${pages}`);
  };

  const exportExcel = () => {
    const targetIds = new Set(exportTargetRows.map((r) => r.w.id));
    const scoped = selected.size > 0;
    const period = view === "month" ? `${MONTHS[m - 1]} ${y}.` : `Godina ${y}.`;
    const subtitle = [period, objFilterName || "svi objekti", exportScopeLabel || null].filter(Boolean).join("  ·  ");
    const cur = (c) => (c === "CZK" ? "Kč" : "€");
    const wName = (id) => data.workers.find((x) => x.id === id)?.name || "";
    const paidStatus = (r) => { if (view !== "month") return ""; const p = paidFor(data, r.w.id, month); return p ? (p.approved ? "Odobreno" : "Predloženo") : "Nije"; };
    const statusColor = (v) => (v === "Odobreno" ? XC.green : v === "Predloženo" ? XC.amber : v === "Nije" ? XC.red : null);
    const net = { bold: true, color: () => XC.amber, fill: () => XC.amberSoft };
    const objRows = [...byObject.entries()].sort(([ka], [kb]) => objName(ka).localeCompare(objName(kb), "hr")).map(([k, v]) => ({
      obj: k === "__none__" ? "Bez objekta" : objName(k), hours: v.hours, grossE: v.gross, grossK: v.grossKc || 0,
      costE: v.costs || 0, costK: v.costsKc || 0, revE: v.revenue, revK: v.revenueKc || 0,
      profE: round2(v.revenue - v.gross - (v.costs || 0)), profK: round2((v.revenueKc || 0) - (v.grossKc || 0) - (v.costsKc || 0)),
    }));
    downloadStyledXlsx({
      fileName: `Obracun_${view === "month" ? MONTHS[m - 1] + "_" + y : "Godina_" + y}${objFilterName ? "_" + objFilterName.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^\w]+/g, "") : ""}.xlsx`,
      sheets: [
        { name: "Obračun", title: `Obračun plaća — ${period}`, subtitle, landscape: true, freezeCols: 1,
          total: { keys: ["hours", "grossE", "grossK", "extraE", "extraK", "bonus", "adv", "ded", "bankE", "bankK", "netE", "netK", "firmE", "firmK"] },
          columns: [
            { key: "w", header: "Radnik", width: 24, bold: true }, { key: "obj", header: "Objekt", width: 16 },
            { key: "hours", header: "Sati", width: 8, type: "hours" }, { key: "rate", header: "Satnica", width: 9, type: "money" },
            { key: "rcur", header: "Valuta", width: 7, type: "center" },
            { key: "grossE", header: "Zarada (€)", width: 11, type: "eur" }, { key: "grossK", header: "Zarada (Kč)", width: 11, type: "czk" },
            { key: "extraE", header: "Vikend / praznik (€)", width: 11, type: "eur" }, { key: "extraK", header: "Vikend / praznik (Kč)", width: 11, type: "czk" },
            { key: "bonus", header: "Bonus (€)", width: 10, type: "eur" }, { key: "adv", header: "Avans (€)", width: 10, type: "eur" },
            { key: "ded", header: "Odbici (€)", width: 10, type: "eur" },
            { key: "bankE", header: "Na račun (€)", width: 11, type: "eur" }, { key: "bankK", header: "Na račun (Kč)", width: 11, type: "czk" },
            { key: "netE", header: "Za isplatu kovertom (€)", width: 13, type: "eur", ...net }, { key: "netK", header: "Za isplatu kovertom (Kč)", width: 13, type: "czk", ...net },
            { key: "firmE", header: "Trošak firme (€)", width: 12, type: "eur" }, { key: "firmK", header: "Trošak firme (Kč)", width: 12, type: "czk" },
            { key: "method", header: "Način isplate", width: 16 },
            ...(view === "month" ? [{ key: "paid", header: "Isplaćeno", width: 12, type: "center", bold: true, color: statusColor }] : []),
          ],
          rows: exportTargetRows.map((r) => ({
            w: r.w.name, obj: objName(r.w.objectId) || "", hours: r.hours,
            rate: r.rateSet.length === 1 ? r.rateSet[0] : r.rateSet.length === 0 ? rateNow(data, r.w) : "razne", rcur: cur(wCur(r.w)),
            grossE: r.gross, grossK: r.grossKc, extraE: r.extraPay, extraK: r.extraPayKc, bonus: r.bonuses, adv: r.advances, ded: r.deductions,
            bankE: r.bank, bankK: r.czk.bank, netE: r.net, netK: r.czk.net, firmE: r.firmCosts, firmK: r.czk.firmCosts,
            method: paymentMethodLabel(r).replace(/[💶🏦+]/g, "").trim(), paid: paidStatus(r),
          })),
          emptyText: "Nema radnika za ovo razdoblje.",
          note: "Zbroj (UKUPNO) prati filter: kad u zaglavlju filtriraš npr. jedan objekt, zbroj se odnosi samo na prikazane redove." },
        { name: "Sati", title: `Sati — ${period}`, subtitle, landscape: true, total: { keys: ["hours"] },
          columns: [
            { key: "date", header: "Datum", width: 12, type: "date" }, { key: "w", header: "Radnik", width: 24, bold: true },
            { key: "obj", header: "Objekt", width: 18 }, { key: "from", header: "Od", width: 8, type: "center" }, { key: "to", header: "Do", width: 8, type: "center" },
            { key: "hours", header: "Sati", width: 8, type: "hours" }, { key: "rate", header: "Satnica", width: 9, type: "money" },
            { key: "amount", header: "Iznos", width: 11, type: "money" }, { key: "cur", header: "Valuta", width: 7, type: "center" },
          ],
          rows: periodLogs.filter((l) => !scoped || targetIds.has(l.workerId)).sort((a, b) => a.date.localeCompare(b.date) || wName(a.workerId).localeCompare(wName(b.workerId), "hr")).map((l) => {
            const wk = data.workers.find((x) => x.id === l.workerId);
            const rate = rateFor(data, wk, l.date);
            return { date: l.date, w: wk?.name || "", obj: objName(l.objectId) || "", from: l.from || "", to: l.to || "", hours: l.hours, rate, amount: round2(l.hours * rate), cur: cur(wCur(wk)) };
          }),
          emptyText: "Nema upisanih sati za ovo razdoblje." },
        { name: "Isplate i troškovi", title: `Isplate i troškovi — ${period}`, subtitle, landscape: true,
          columns: [
            { key: "date", header: "Datum", width: 12, type: "date" }, { key: "za", header: "Za", width: 26, bold: true },
            { key: "type", header: "Vrsta", width: 14 }, { key: "amount", header: "Iznos", width: 12, type: "money" },
            { key: "cur", header: "Valuta", width: 7, type: "center" }, { key: "note", header: "Napomena", width: 36, wrap: true },
          ],
          rows: data.payments.filter((pp) => (view === "month" ? inMonth : inYear)(pp.date) && (!scoped || !pp.workerId || targetIds.has(pp.workerId))).sort((a, b) => a.date.localeCompare(b.date)).map((pp) => ({
            date: pp.date, za: pp.workerId ? wName(pp.workerId) : "Objekt: " + (objName(pp.objectId) || ""),
            type: TYPE_LABEL[pp.type] || pp.type, amount: pp.amount, cur: cur(pp.currency), note: pp.note || "",
          })),
          emptyText: "Nema isplata ni troškova za ovo razdoblje." },
        { name: "Po objektima", title: `Po objektima — ${period}`, subtitle, landscape: true,
          total: { keys: ["hours", "grossE", "grossK", "costE", "costK", ...(admin ? ["revE", "revK", "profE", "profK"] : [])] },
          columns: [
            { key: "obj", header: "Objekt", width: 22, bold: true }, { key: "hours", header: "Sati", width: 9, type: "hours" },
            { key: "grossE", header: "Zarada radnika (€)", width: 13, type: "eur" }, { key: "grossK", header: "Zarada radnika (Kč)", width: 13, type: "czk" },
            { key: "costE", header: "Troškovi objekta (€)", width: 13, type: "eur" }, { key: "costK", header: "Troškovi objekta (Kč)", width: 13, type: "czk" },
            ...(admin ? [
              { key: "revE", header: "Naplata (€)", width: 13, type: "eur" }, { key: "revK", header: "Naplata (Kč)", width: 13, type: "czk" },
              { key: "profE", header: "Dobit (€)", width: 13, type: "eur", bold: true, color: (v) => (v < 0 ? XC.red : XC.green) },
              { key: "profK", header: "Dobit (Kč)", width: 13, type: "czk", bold: true, color: (v) => (v < 0 ? XC.red : XC.green) },
            ] : []),
          ],
          rows: objRows, emptyText: "Nema podataka po objektima." },
      ],
    }).catch((e) => alert("Izvoz nije uspio: " + (e.message || e)));
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
        <div style={{ fontWeight: 800, fontSize: 16.5 }}>{periodLabelUi}</div>
        <Btn small kind="ghost" onClick={view === "month" ? next : () => setMonth(`${y + 1}-${String(m).padStart(2, "0")}`)}>→</Btn>
      </div>

      <div style={{ marginBottom: 12 }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: S.sub, marginBottom: 5 }}>{api.t("showFor")}</div>
        <div onClick={() => setObjPickerOpen(!objPickerOpen)} style={{
          border: `1px solid ${S.line}`, borderRadius: 10, padding: "10px 12px", cursor: "pointer", background: "#fff",
          display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
          <span>{objFilterLabel}</span>
          <span style={{ color: S.sub, fontWeight: 700 }}>{objPickerOpen ? "▲" : "▼"}</span>
        </div>
        {objPickerOpen && (
          <div style={{ border: `1px solid ${S.line}`, borderRadius: 10, marginTop: 6, padding: "6px 12px", maxHeight: 280, overflowY: "auto" }}>
            <label style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", fontWeight: 700, borderBottom: `1px solid ${S.line}`, cursor: "pointer" }}>
              <input type="checkbox" checked={objFilters.size === 0} onChange={() => setObjFilters(new Set())} style={{ width: 17, height: 17 }} />
              {api.t("allObjects")}
            </label>
            {["HR", "CZ"].map((c) => {
              const objs = sortedObjects(data.objects.filter((o) => (o.country || "HR") === c));
              if (!objs.length) return null;
              return (
                <div key={c}>
                  <div style={{ fontSize: 11.5, fontWeight: 700, color: S.sub, margin: "8px 0 2px" }}>{FLAG[c]} {COUNTRY_NAME[c]}</div>
                  {objs.map((o) => (
                    <label key={o.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 0", fontSize: 13.5, cursor: "pointer" }}>
                      <input type="checkbox" checked={objFilters.has(o.id)} onChange={() => toggleObjFilter(o.id)} style={{ width: 17, height: 17 }} />
                      {o.name}
                    </label>
                  ))}
                </div>
              );
            })}
          </div>
        )}
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap", alignItems: "flex-start" }}>
        {rows.length > 0 && (
          <div style={{ position: "relative" }}>
            <Btn small kind="ghost" onClick={() => setExportOpen(!exportOpen)} style={{ fontWeight: 700 }}>{tr("Izvoz i ispis")} {exportOpen ? "▲" : "▼"}</Btn>
            {exportOpen && (
              <>
                <div onClick={() => setExportOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 29 }} />
                <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, zIndex: 30, background: "#fff", border: `1px solid ${S.line}`,
                  borderRadius: 12, boxShadow: "0 8px 24px rgba(28,37,33,.12)", minWidth: 250, padding: 6 }}>
                  <div style={{ fontSize: 11.5, color: S.sub, padding: "4px 10px 6px" }}>{exportScopeText}</div>
                  {[[tr("Excel tablica"), exportExcel], [tr("Koverte za isplatu"), () => printEnvelopes(exportTargetRows)], [tr("PDF obračuni radnika"), printAllPayslips]].map(([label, fn]) => (
                    <button key={label} onClick={() => { setExportOpen(false); fn(); }} style={{ display: "block", width: "100%", textAlign: "left", background: "none", border: "none",
                      borderRadius: 8, padding: "9px 10px", fontSize: 14, fontWeight: 600, color: S.ink, cursor: "pointer" }}>{label}</button>
                  ))}
                </div>
              </>
            )}
          </div>
        )}
        {admin && view === "month" && (
          <Btn small kind="ghost" onClick={() => mzdyFileRef.current?.click()} disabled={mzdyBusy}>
            {mzdyBusy ? tr("Čitam datoteke…") : tr("Učitaj Mzdy (PDF od knjigovođe)")}
          </Btn>
        )}
        <input ref={mzdyFileRef} type="file" accept="application/pdf" multiple onChange={handleMzdyFile} style={{ display: "none" }} />
        {mzdyErr && <div style={{ color: S.red, fontSize: 12.5, width: "100%" }}>{mzdyErr}</div>}
      </div>

      {mzdyReview && (() => {
        const setRow = (i, patch) => { const rows = [...mzdyReview.rows]; rows[i] = { ...rows[i], ...patch }; setMzdyReview({ ...mzdyReview, rows }); };
        const chosen = mzdyReview.rows.filter((r) => r.include && r.workerId && !r.locked);
        const sumBy = (k) => chosen.reduce((s, r) => s + (r.currency === "CZK" ? (parseNum(r[k]) || 0) : 0), 0);
        return (
        <div style={{ position: "fixed", inset: 0, background: "rgba(10,14,12,.92)", zIndex: 60, overflowY: "auto", padding: "20px 12px 40px" }}>
          <div style={{ maxWidth: 560, margin: "0 auto", background: "#fff", borderRadius: 14, padding: 18 }}>
            <div style={{ fontWeight: 800, fontSize: 17, marginBottom: 4 }}>{tr("📄 Mzdy — provjera prije spremanja")}</div>
            <div style={{ fontSize: 12.5, color: S.sub, marginBottom: 10 }}>
              {tr("Na kovertu se oduzima")}{" "}<b>{tr("soc.-zdr.")}</b> {tr("(puno radno vrijeme")}{" "}{czk(SOCZDR_FULL)}{tr(", pola")}{" "}{czk(SOCZDR_HALF)}{tr(", DPČ 0) i")}{" "}<b>{tr("na račun")}</b> {tr("(Dobírka iz pásky). Svaki iznos možeš ispraviti.")}
            </div>
            <div style={{ background: S.bg, borderRadius: 10, padding: "8px 10px", fontSize: 12, color: S.sub, marginBottom: 12 }}>
              {mzdyReview.docs.map((d, i) => <div key={i}>✓ {d}</div>)}
            </div>
            <Field label={tr("Mjesec")}>
              <input type="month" value={mzdyReview.month} onChange={(e) => setMzdyReview({ ...mzdyReview, month: e.target.value })} />
            </Field>
            {mzdyReview.rows.map((row, i) => (
              <div key={i} style={{ border: `1px solid ${row.include && row.workerId ? S.green : S.line}`, borderRadius: 12, padding: "10px 12px", marginBottom: 10, opacity: row.locked ? 0.6 : 1 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <input type="checkbox" checked={row.include} disabled={row.locked || !row.workerId} onChange={(e) => setRow(i, { include: e.target.checked })} style={{ width: 18, height: 18 }} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 700, fontSize: 14.5 }}>{row.rawName} {row.corrected && <Tag color={S.blue} bg={S.blueSoft}>{tr("oprava")}</Tag>}</div>
                    <div className="num" style={{ fontSize: 12, color: S.sub }}>
                      {row.kat}{row.tarif ? tr(" · tarif {0} ({1} radno vrijeme)", czk(row.tarif), row.tarif >= FULL_TARIF ? tr("puno") : tr("pola")) : ""}
                    </div>
                  </div>
                </div>
                <Field label={tr("Radnik u aplikaciji")}>
                  <select value={row.workerId} onChange={(e) => {
                    const wk = data.workers.find((x) => x.id === e.target.value);
                    const cur = wk ? wCur(wk) : "CZK";
                    const toCur = (v) => String(cur === row.currency ? v : cur === "CZK" ? round2((parseNum(v) || 0) * czkRate) : round2((parseNum(v) || 0) / czkRate));
                    setRow(i, { workerId: e.target.value, currency: cur, socZdr: toCur(row.socZdr), bank: toCur(row.bank), include: !!e.target.value && !row.locked });
                  }}>
                    <option value="">{tr("— nije prepoznat, odaberi ručno —")}</option>
                    {sortedWorkers(data.workers).map((w) => <option key={w.id} value={w.id}>{w.name}{w.archived ? tr(" (bivši)") : ""}</option>)}
                  </select>
                </Field>
                <div style={{ display: "flex", gap: 8 }}>
                  <div style={{ flex: 1 }}>
                    <Field label={tr("soc.-zdr. ({0})", row.currency === "CZK" ? "Kč" : "€")}>
                      <input inputMode="decimal" value={row.socZdr} onChange={(e) => setRow(i, { socZdr: e.target.value })} />
                    </Field>
                  </div>
                  <div style={{ flex: 1 }}>
                    <Field label={tr("Na račun ({0})", row.currency === "CZK" ? "Kč" : "€")}>
                      <input inputMode="decimal" value={row.bank} disabled={!row.addBank} onChange={(e) => setRow(i, { bank: e.target.value })} />
                    </Field>
                  </div>
                </div>
                <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12.5, cursor: "pointer" }}>
                  <input type="checkbox" checked={row.addBank} onChange={(e) => setRow(i, { addBank: e.target.checked })} style={{ width: 16, height: 16 }} />
                  {tr("Upiši \"na račun\" iz pásky")}
                </label>
                {row.warn.map((w, k) => <div key={k} style={{ fontSize: 12, color: S.amber, marginTop: 4 }}>⚠️ {w}</div>)}
              </div>
            ))}
            {mzdyReview.prikazy.length > 0 && (
              <div style={{ border: `1px solid ${S.line}`, borderRadius: 12, padding: "10px 12px", marginBottom: 10 }}>
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 6 }}>{tr("🏦 Příkazy k úhradě")}</div>
                {mzdyReview.prikazy.map((p, k) => (
                  <div key={k} className="num" style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "3px 0" }}>
                    <span>{p.to} · {p.label}</span><b>{czk(p.amount)}</b>
                  </div>
                ))}
                <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12.5, cursor: "pointer", marginTop: 6 }}>
                  <input type="checkbox" checked={mzdyReview.addReminders} onChange={(e) => setMzdyReview({ ...mzdyReview, addReminders: e.target.checked })} style={{ width: 16, height: 16 }} />
                  {tr("Dodaj podsjetnike za plaćanje (rok 20. idućeg mjeseca)")}
                </label>
              </div>
            )}
            <div className="num" style={{ background: S.greenSoft, borderRadius: 10, padding: "8px 12px", fontSize: 13, marginBottom: 10 }}>
              {tr("Spremit će se za")}{" "}<b>{chosen.length}</b> {tr("radnika: soc.-zdr.")}{" "}<b>{czk(sumBy("socZdr"))}</b> {tr("· na račun")}{" "}<b>{czk(chosen.filter((r) => r.addBank).reduce((s, r) => s + (r.currency === "CZK" ? (parseNum(r.bank) || 0) : 0), 0))}</b>
            </div>
            {mzdyErr && (
              <div style={{ background: S.redSoft, color: S.red, borderRadius: 8, padding: "8px 10px", fontSize: 12.5, fontWeight: 600, marginBottom: 8 }}>{mzdyErr}</div>
            )}
            <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
              <Btn onClick={saveMzdyReview} disabled={mzdyBusy} style={{ flex: 1 }}>{mzdyBusy ? tr("Spremam…") : tr("💾 Spremi")}</Btn>
              <Btn kind="ghost" onClick={() => { setMzdyReview(null); setMzdyErr(""); }}>{tr("Odustani")}</Btn>
            </div>
          </div>
        </div>
        );
      })()}

      {rows.length === 0 ? <Empty text={objFilters.size > 0 ? tr("Za {0} u ovom razdoblju nema upisanih sati.", objFilterName) : tr("Za ovo razdoblje nema upisanih sati ni isplata.")} /> : (
        <>
          {(totals.grossKc !== 0 || totals.netKc !== 0 || totals.firmKc !== 0) && (
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: S.sub, cursor: "pointer", justifyContent: "flex-end", margin: "0 2px 6px" }}>
              <input type="checkbox" checked={showCombinedEur} onChange={(e) => setShowCombinedEur(e.target.checked)} style={{ width: 15, height: 15 }} />
              {tr("Prikaži i zbroj u € (tečaj")}{" "}{czkRate})
            </label>
          )}
          <Card style={{ background: S.ink, color: "#fff", border: "none" }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, fontSize: 13.5 }}>
              <div><div style={{ color: "#9DB3A8" }}>{api.t("totalHours")}</div><div className="num" style={{ fontSize: 18, fontWeight: 800 }}>{fmtH(totals.hours)}</div></div>
              <div><div style={{ color: "#9DB3A8" }}>{api.t("earningsBonus")}</div>
                <TwoCur e={round2(totals.gross + totals.bonus)} k={round2(totals.grossKc + totals.bonusKc)} />
                {showCombinedEur && round2(totals.grossKc + totals.bonusKc) !== 0 && <div className="num" style={{ fontSize: 11.5, color: "#9DB3A8" }}>≈ {eur(round2(totals.gross + totals.bonus + (totals.grossKc + totals.bonusKc) / czkRate))} {tr("ukupno")}</div>}
              </div>
              <div><div style={{ color: "#9DB3A8" }}>{api.t("toPay")}</div>
                <TwoCur e={totals.net} k={totals.netKc} color="#7FD6B4" />
                {showCombinedEur && totals.netKc !== 0 && <div className="num" style={{ fontSize: 11.5, color: "#9DB3A8" }}>≈ {eur(round2(totals.net + totals.netKc / czkRate))} {tr("ukupno")}</div>}
              </div>
              <div><div style={{ color: "#9DB3A8" }}>{api.t("firmCosts")}</div>
                <TwoCur e={round2(totals.firm + objCostsTotal)} k={round2(totals.firmKc + objCostsTotalKc)} />
                {showCombinedEur && round2(totals.firmKc + objCostsTotalKc) !== 0 && <div className="num" style={{ fontSize: 11.5, color: "#9DB3A8" }}>≈ {eur(round2(totals.firm + objCostsTotal + (totals.firmKc + objCostsTotalKc) / czkRate))} {tr("ukupno")}</div>}
              </div>
            </div>
          </Card>

          <div style={{ display: "flex", background: "#fff", border: `1px solid ${S.line}`, borderRadius: 12, padding: 3, marginBottom: 12 }}>
            {[["radnici", tr("Radnici"), displayRows.length], ["objekti", tr("Objekti"), [...byObject.keys()].filter((k) => k !== "__none__").length], ["analiza", admin ? tr("Dobit i analiza") : tr("Analiza"), null]].map(([id, label, n]) => (
              <button key={id} onClick={() => setSection(id)} style={{ flex: 1, padding: "8px 6px", borderRadius: 9, border: "none", fontWeight: 700, fontSize: 13.5, cursor: "pointer",
                background: section === id ? S.ink : "transparent", color: section === id ? "#fff" : S.sub }}>
                {label}{n != null ? <span style={{ opacity: 0.7, fontWeight: 600 }}> {n}</span> : null}
              </button>
            ))}
          </div>

          {section === "radnici" && (
            <>
              {view === "month" && (
                <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
                  {[["all", "Svi"], ["paid", "✓ Isplaćeno"], ["pending", "⏳ Predloženo"], ["unpaid", "Bez isplate"]].map(([id, label]) => (
                    <button key={id} onClick={() => setPayStatusFilter(id)} style={{ padding: "6px 12px", borderRadius: 999, fontWeight: 700, fontSize: 12.5, cursor: "pointer",
                      background: payStatusFilter === id ? S.ink : "#fff", color: payStatusFilter === id ? "#fff" : S.sub, border: `1px solid ${payStatusFilter === id ? S.ink : S.line}` }}>{tr(label)}</button>
                  ))}
                </div>
              )}
              {view === "month" && (unpaidRows.length > 0 || pendingRows.length > 0) && (
                <div style={{ display: "flex", alignItems: "center", gap: 14, padding: "2px 4px 10px", flexWrap: "wrap" }}>
                  {unpaidRows.length > 0 && (
                    <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, cursor: "pointer" }}>
                      <input type="checkbox" checked={allSelected} onChange={toggleSelectAll} style={{ width: 18, height: 18 }} />
                      {tr("Označi neisplaćene ({0})", unpaidRows.length)}
                    </label>
                  )}
                  {pendingRows.length > 0 && (
                    <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, cursor: "pointer" }}>
                      <input type="checkbox" checked={allPendingSelected} onChange={toggleSelectAllPending} style={{ width: 18, height: 18 }} />
                      {tr("Označi predložene ({0})", pendingRows.length)}
                    </label>
                  )}
                </div>
              )}

              {displayRows.length === 0 && <Empty text={tr("Nitko ne odgovara ovom filteru.")} />}
              {displayRows.map((r) => {
                const paid = view === "month" ? paidFor(data, r.w.id, month) : null;
                const gc = wCur(r.w);
                const bdHead = breakdownOf(r); // isti iznos kao "Za isplatu" u razradi i na koverti (uključuje soc.-zdr.)
                const main = gc === "CZK" ? bdHead.netK : bdHead.netE;
                const other = gc === "CZK" ? bdHead.netE : bdHead.netK;
                return (
                  <Card key={r.w.id} style={paid ? { borderColor: "#C5DED2", background: "#FBFDF9" } : undefined}>
                    <div onClick={() => setOpen(open === r.w.id ? null : r.w.id)} style={{ cursor: "pointer" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                        <div style={{ fontWeight: 700, fontSize: 15.5, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", minWidth: 0 }}>
                          <input type="checkbox" checked={selected.has(r.w.id)} onChange={(e) => toggleSelect(r.w.id, e)} onClick={(e) => e.stopPropagation()} style={{ width: 18, height: 18 }} title={tr("Označi za izvoz ili grupnu radnju")} />
                          <span onClick={(e) => { if (onOpenWorker) { e.stopPropagation(); onOpenWorker(r.w.id); } }} style={{ color: onOpenWorker ? S.blue : "inherit" }}>{r.w.name}</span>
                          {paid && (paid.approved
                            ? <Tag color={S.green} bg={S.greenSoft}>{tr("✓ Isplaćeno")}</Tag>
                            : <Tag color={S.amber} bg={S.amberSoft}>{tr("⏳ Predloženo")}</Tag>)}
                        </div>
                        <div className="num" style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                          <div style={{ fontWeight: 800, color: main >= 0 ? S.green : S.red, fontSize: 16 }}>{money(main, gc)}</div>
                          {other !== 0 && <div style={{ fontWeight: 700, color: other >= 0 ? S.green : S.red, fontSize: 12.5 }}>{other > 0 ? "+" : "−"}{money(Math.abs(other), gc === "CZK" ? "EUR" : "CZK")}</div>}
                        </div>
                      </div>
                      <div className="num" style={{ fontSize: 13, color: S.sub, marginTop: 3 }}>
                        {view === "month" && isHoursSuspicious(r.hours) && <span title={tr("Više od {0}h u mjesecu — provjeri unos", MONTH_HOURS_WARN)} style={{ color: S.red }}>⚠️ </span>}
                        {fmtH(r.hours)}{r.rateSet.length === 1 ? <> × {money(r.rateSet[0], gc)}</> : r.rateSet.length > 1 ? <> {tr("(više satnica)")}</> : null} = {gc === "CZK" ? czk(r.grossKc) : eur(r.gross)}
                        {open !== r.w.id && bdHead.items.length > 1 && <span> · {tr("otvori za razradu")}</span>}
                      </div>
                    </div>
                    {open === r.w.id && (() => {
                      const bd = breakdownOf(r);
                      const head = { fontSize: 11.5, fontWeight: 700, color: S.sub, textTransform: "uppercase", letterSpacing: 0.5, margin: "12px 0 4px" };
                      return (
                      <div style={{ marginTop: 10, borderTop: `1px solid ${S.line}`, paddingTop: 2 }}>
                        <div style={head}>{tr("Sati")}</div>
                        {r.logs.sort((a, b) => a.date.localeCompare(b.date)).map((l) => (
                          <div key={l.id} className="num" style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13.5, padding: "4px 0" }}>
                            <span style={{ flex: 1 }}>{fmtDate(l.date)} · <LogSpanLabel l={l} />{objName(l.objectId) ? " · " + objName(l.objectId) : ""}</span>
                            <span style={{ fontWeight: 600 }}>{fmtH(l.hours)}</span>
                            {(admin || l.createdBy === api.uid()) && (
                              <button onClick={(e) => { e.stopPropagation(); api.delLog(l, r.w.name); }} title={tr("Obriši")}
                                style={{ background: "none", border: "none", color: S.red, fontSize: 15, cursor: "pointer", padding: "0 2px" }}>✕</button>
                            )}
                          </div>
                        ))}
                        {(r.pays.length > 0 || bd.note) && <div style={head}>{tr("Isplate i odbici")}</div>}
                        {r.pays.map((p) => (
                          <div key={p.id} className="num" style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13.5, padding: "4px 0",
                            color: p.type === "bonus" ? S.green : p.type === "racun" ? S.blue : p.deduct ? S.amber : S.sub }}>
                            <span style={{ flex: 1 }}>{fmtDate(p.date)} · {tr(TYPE_LABEL[p.type] || p.type)}{p.note ? " · " + p.note : ""}</span>
                            <span style={{ fontWeight: 600 }}>{p.type === "bonus" ? "+" : p.deduct ? "−" : ""}{money(p.amount, p.currency)}</span>
                            <button onClick={(e) => { e.stopPropagation(); api.delPayment(p, r.w.name); }} title={tr("Obriši")}
                              style={{ background: "none", border: "none", color: S.red, fontSize: 15, cursor: "pointer", padding: "0 2px" }}>✕</button>
                          </div>
                        ))}
                        {bd.note && (
                          <div className="num" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13.5, padding: "4px 0", color: S.blue }}>
                            <span>{bd.note.note || MZDY_NOTE_LABEL} {tr("(pola zbroja,")}{" "}{czk(bd.note.sourceCzk)})</span>
                            <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                              <span style={{ fontWeight: 600 }}>−{money(bd.note.amount, bd.note.currency)}</span>
                              <button onClick={() => api.delPayrollNote(bd.note, r.w.name)} style={{ background: "none", border: "none", color: S.red, fontSize: 15, cursor: "pointer", padding: 2 }}>✕</button>
                            </span>
                          </div>
                        )}
                        <div style={head}>{tr("Obračun")}</div>
                        <div className="num" style={{ background: S.bg, borderRadius: 10, padding: "8px 12px", fontSize: 13.5 }}>
                          {bd.items.map((it, i) => (
                            <div key={i} style={{ display: "flex", justifyContent: "space-between", gap: 10, padding: "3px 0" }}>
                              <span style={{ color: S.sub }}>{tr(it.label)}{it.hint ? ` (${it.hint})` : ""}</span>
                              <span style={{ fontWeight: 600 }}>{it.sign}{amtText(it.e, it.k, bd.gc)}</span>
                            </div>
                          ))}
                          <div style={{ borderTop: `1px solid ${S.line}`, marginTop: 6, paddingTop: 6, display: "flex", justifyContent: "space-between", fontWeight: 800, fontSize: 15 }}>
                            <span>{tr("Za isplatu")}</span>
                            <span style={{ color: bd.netE < 0 || bd.netK < 0 ? S.red : S.green }}>{bd.netText}</span>
                          </div>
                        </div>
                        <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
                          <Btn small kind="ghost" onClick={() => printPayslip(r)}>{tr("PDF obračun")}</Btn>
                          <Btn small kind="ghost" onClick={() => shareSummary(r)}>{tr("Podijeli sažetak")}</Btn>
                          <Btn small kind="ghost" onClick={() => setSigTarget({ id: r.w.id, name: r.w.name })} style={signatures[r.w.id] ? { color: S.green, borderColor: "#C5DED2" } : undefined}>
                            {signatures[r.w.id] ? tr("✓ Potpisano (promijeni)") : tr("Potpis radnika")}
                          </Btn>
                          {view === "month" && !paid && (
                            <Btn small onClick={() => api.markPaid(r.w, month, r.net, r.czk.net)}>{tr("✓ Predloži isplatu")}</Btn>
                          )}
                          {view === "month" && paid && !paid.approved && (
                            <>
                              <span className="num" style={{ fontSize: 12.5, color: S.amber, alignSelf: "center", fontWeight: 700 }}>
                                {tr("⏳ Predloženo")}{" "}{fmtDate(paid.paidAt)} · {[paid.amount ? eur(paid.amount) : "", paid.amountKc ? czk(paid.amountKc) : ""].filter(Boolean).join(" + ") || eur(0)}
                              </span>
                              {admin && <Btn small onClick={() => api.approvePayout(paid, r.w.name)}>{tr("✓ Odobri")}</Btn>}
                              <Btn small kind="danger" onClick={() => api.unmarkPaid(paid, r.w.name)}>{tr("Odbaci")}</Btn>
                            </>
                          )}
                          {view === "month" && paid && !paid.approved && (round2(paid.amount || 0) !== round2(r.net || 0) || round2(paid.amountKc || 0) !== round2(r.czk.net || 0)) && (
                            <div style={{ width: "100%", display: "flex", alignItems: "center", gap: 8, background: S.redSoft, borderRadius: 8, padding: "6px 10px" }}>
                              <span style={{ fontSize: 12, color: S.red, fontWeight: 700, flex: 1 }}>
                                {tr("⚠️ Sati/isplate su se promijenili nakon prijedloga — sad bi iznos bio")}{" "}{[r.net ? eur(r.net) : "", r.czk.net ? czk(r.czk.net) : ""].filter(Boolean).join(" + ") || eur(0)}
                              </span>
                              <Btn small onClick={() => api.refreshPayout(paid, r.net, r.czk.net, r.w.name)}>{tr("🔄 Osvježi")}</Btn>
                            </div>
                          )}
                          {view === "month" && paid && paid.approved && (
                            <>
                              <span className="num" style={{ fontSize: 12.5, color: S.green, alignSelf: "center", fontWeight: 700 }}>
                                {tr("✓ Odobreno")}{" "}{fmtDate(paid.approvedAt || paid.paidAt)} · {[paid.amount ? eur(paid.amount) : "", paid.amountKc ? czk(paid.amountKc) : ""].filter(Boolean).join(" + ") || eur(0)}
                              </span>
                              {admin && <Btn small kind="danger" onClick={() => api.unmarkPaid(paid, r.w.name)}>{tr("Otključaj")}</Btn>}
                            </>
                          )}
                        </div>
                        {view === "month" && paid && !paid.approved && (
                          <div style={{ fontSize: 11.5, color: S.sub, marginTop: 6 }}>{tr("Isplata je predložena, ali još nije odobrena — sati se mogu ispraviti do odobrenja.")}</div>
                        )}
                        {view === "month" && paid && paid.approved && (
                          <div style={{ fontSize: 11.5, color: S.sub, marginTop: 6 }}>{tr("Mjesec je odobren i zaključan — sati i isplate se više ne mogu mijenjati")}{admin ? "" : tr(" (otključava admin)")}.</div>
                        )}
                      </div>
                      );
                    })()}
                  </Card>
                );
              })}

              {selected.size > 0 && (() => {
                const selectedUnpaid = unpaidRows.filter((r) => selected.has(r.w.id));
                const selectedPending = pendingRows.filter((r) => selected.has(r.w.id));
                return (
                  <div style={{ position: "sticky", bottom: "calc(72px + env(safe-area-inset-bottom))", zIndex: 15, background: S.ink, color: "#fff", borderRadius: 14,
                    padding: "10px 12px", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", boxShadow: "0 8px 24px rgba(28,37,33,.25)", marginTop: 4 }}>
                    <span style={{ fontWeight: 700, fontSize: 13.5 }}>{tr("Odabrano: {0}", selected.size)}</span>
                    <div style={{ marginLeft: "auto", display: "flex", gap: 8, flexWrap: "wrap" }}>
                      {view === "month" && selectedUnpaid.length > 0 && (
                        <>
                          <Btn small kind="ghost" onClick={() => payMany(selectedUnpaid)} disabled={bulkBusy}>{tr("Predloži isplatu ({0})", selectedUnpaid.length)}</Btn>
                          <Btn small onClick={() => payAndPrintMany(selectedUnpaid)} disabled={bulkBusy}>{bulkBusy ? tr("Spremam…") : tr("Predloži i ispiši koverte ({0})", selectedUnpaid.length)}</Btn>
                        </>
                      )}
                      {view === "month" && admin && selectedPending.length > 0 && (
                        <Btn small onClick={() => approveMany(selectedPending)} disabled={bulkBusy}>{bulkBusy ? tr("Spremam…") : tr("Odobri ({0})", selectedPending.length)}</Btn>
                      )}
                      <Btn small kind="ghost" onClick={() => setSelected(new Set())} disabled={bulkBusy}>{tr("Poništi odabir")}</Btn>
                    </div>
                  </div>
                );
              })()}
            </>
          )}

          {section === "objekti" && (byObject.size === 0 ? <Empty text={tr("Nema podataka po objektima.")} /> : (
            <>
              {/* po objektima + naplata od hotela */}
              {byObject.size > 0 && (
                <Card>
                  <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("Po objektima")}</div>
                  {["HR", "CZ", "__none__"].map((cGroup) => {
                    const entries = [...byObject.entries()].filter(([key]) => {
                      if (cGroup === "__none__") return key === "__none__";
                      const ob = data.objects.find((o) => o.id === key);
                      return ob && (ob.country || "HR") === cGroup;
                    }).sort(([ka], [kb]) => objName(ka).localeCompare(objName(kb), "hr"));
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
                          <div className="num" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0 3px", fontWeight: 800, fontSize: 13.5 }}>
                            <span>{FLAG[cGroup]} {COUNTRY_NAME[cGroup]}</span>
                            <span>{fmtH(sub.hours)}{admin ? <> {tr("· dobit")}{" "}{[sub.revenue || sub.gross || sub.costs ? eur(subProfitE) : "", sub.revenueKc || sub.grossKc || sub.costsKc ? czk(subProfitK) : ""].filter(Boolean).join(" · ") || eur(0)}</> : null}</span>
                          </div>
                        )}
                        {entries.map(([key, v]) => {
                    const ups = view === "month" ? (data.invoicePayments || []).filter((p) => p.objectId === key && p.month === month) : [];
                    const paid = round2(ups.reduce((s, p) => s + p.amount, 0));
                    const bcKey = data.objects.find((o) => o.id === key)?.billCur || "EUR";
                    const due = round2((bcKey === "CZK" ? v.revenueKc : v.revenue) - paid);
                    return (
                      <div key={key} style={{ padding: "7px 0", borderBottom: `1px solid ${S.line}` }}>
                        <div className="num" style={{ display: "flex", justifyContent: "space-between", fontSize: 14 }}>
                          <span style={{ fontWeight: 600 }}>{key === "__none__" ? tr("Bez objekta") : objName(key)}</span>
                          <span>{fmtH(v.hours)} · <b>{[v.gross ? eur(v.gross) : "", v.grossKc ? czk(v.grossKc) : ""].filter(Boolean).join(" + ") || eur(0)}</b>{(v.costs > 0 || v.costsKc > 0) ? <span style={{ color: S.red }}> {tr("· troškovi")}{" "}{[v.costs ? eur(v.costs) : "", v.costsKc ? czk(v.costsKc) : ""].filter(Boolean).join(" + ")}</span> : null}</span>
                        </div>
                        {admin && key !== "__none__" && (
                          <>
                            <div className="num" style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, color: S.sub, fontWeight: 700 }}>
                              <span>{tr("naplata")}{" "}{[v.revenue ? eur(v.revenue) : "", v.revenueKc ? czk(v.revenueKc) : ""].filter(Boolean).join(" + ") || eur(0)}</span>
                              <span>{tr("dobit")}{" "}{[
                                (v.revenue || v.gross || v.costs) ? eur(round2(v.revenue - v.gross - (v.costs || 0))) : "",
                                (v.revenueKc || v.grossKc || v.costsKc) ? czk(round2((v.revenueKc || 0) - (v.grossKc || 0) - (v.costsKc || 0))) : "",
                              ].filter(Boolean).join(" · ") || eur(0)}</span>
                            </div>
                            {view === "month" && (
                              <div style={{ marginTop: 6 }}>
                                <div className="num" style={{ fontSize: 12.5, display: "flex", justifyContent: "space-between" }}>
                                  <span>{tr("uplaćeno:")}{" "}<b style={{ color: S.green }}>{money(paid, bcKey)}</b></span>
                                  <span>{tr("ostaje:")}{" "}<b style={{ color: due > 0 ? S.red : S.green }}>{money(Math.max(due, 0), bcKey)}</b></span>
                                </div>
                                {ups.map((p) => (
                                  <div key={p.id} className="num" style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: S.sub, padding: "2px 0" }}>
                                    <span>{fmtDate(p.date)} {tr("· uplata")}{" "}{money(p.amount, bcKey)}</span>
                                    <button onClick={() => api.delInvoicePayment(p, objName(key))} style={{ background: "none", border: "none", color: S.red, cursor: "pointer" }}>✕</button>
                                  </div>
                                ))}
                                <div style={{ display: "flex", gap: 6, marginTop: 4, flexWrap: "wrap" }}>
                                  <input inputMode="decimal" placeholder={bcKey === "CZK" ? tr("iznos uplate Kč") : tr("iznos uplate €")} value={uplata[key] || ""}
                                    onChange={(e) => setUplata({ ...uplata, [key]: e.target.value })} style={{ padding: "7px 9px", fontSize: 13, flex: "1 1 120px", minWidth: 0 }} />
                                  <Btn small kind="ghost" onClick={() => {
                                    const a = parseNum(uplata[key]);
                                    if (a) { api.addInvoicePayment(data.objects.find((o) => o.id === key), month, round2(a)); setUplata({ ...uplata, [key]: "" }); }
                                  }}>{tr("+ Uplata")}</Btn>
                                  <Btn small kind="ghost" onClick={() => printInvoiceSpec(key, v)}>{tr("Specifikacija PDF")}</Btn>
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
            </>
          ))}

          {section === "analiza" && (
            <>
              {admin && (
                <Card>
                  <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("Dobit ({0})", periodLabelUi)}</div>
                  <div className="num" style={{ fontSize: 14 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Naplata od objekata")}</span><b>{eur(totalRevenue)}{totalRevenueKc !== 0 ? " + " + czk(totalRevenueKc) : ""}</b></div>
                    <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0" }}><span>{tr("Plaće + bonusi + svi troškovi")}</span><b>−{eur(totalCost)}{totalCostKc !== 0 ? " − " + czk(totalCostKc) : ""}</b></div>
                    <div style={{ borderTop: `1px solid ${S.line}`, paddingTop: 6 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 800, fontSize: 15.5 }}>
                        <span>{tr("DOBIT (€)")}</span><span style={{ color: totalProfit >= 0 ? S.green : S.red }}>{eur(totalProfit)}</span>
                      </div>
                      {(totalRevenueKc !== 0 || totalCostKc !== 0) && (
                        <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 800, fontSize: 15.5 }}>
                          <span>{tr("DOBIT (Kč)")}</span><span style={{ color: totalProfitKc >= 0 ? S.green : S.red }}>{czk(totalProfitKc)}</span>
                        </div>
                      )}
                    </div>
                  </div>
                </Card>
              )}
              {/* grafovi */}
              <Card>
                <div style={{ fontWeight: 700, marginBottom: 4 }}>{tr("Sati — zadnjih 6 mjeseci")}</div>
                <MiniBars items={last6.map((x) => ({ label: x.label, value: x.hours }))} fmt={(v) => Math.round(v)} color={S.blue} />
                {admin && (
                  <>
                    <div style={{ fontWeight: 700, margin: "12px 0 4px" }}>{tr("Dobit u € — zadnjih 6 mjeseci")}</div>
                    <MiniBars items={last6.map((x) => ({ label: x.label, value: x.profit }))} fmt={(v) => Math.round(v) + "€"} color={S.green} />
                    {last6.some((x) => x.profitKc !== 0) && (
                      <>
                        <div style={{ fontWeight: 700, margin: "12px 0 4px" }}>{tr("Dobit u Kč — zadnjih 6 mjeseci")}</div>
                        <MiniBars items={last6.map((x) => ({ label: x.label, value: x.profitKc }))} fmt={(v) => Math.round(v) + " Kč"} color={S.green} />
                      </>
                    )}
                  </>
                )}
              </Card>
              {/* rang objekata po dobiti */}
              {admin && byObject.size > 1 && (() => {
                const ranked = [...byObject.entries()]
                  .filter(([key]) => key !== "__none__")
                  .map(([key, v]) => {
                    const profitE = round2(v.revenue - v.gross - (v.costs || 0));
                    const profitK = round2((v.revenueKc || 0) - (v.grossKc || 0) - (v.costsKc || 0));
                    return { key, name: objName(key), profitE, profitK, combined: round2(profitE + profitK / czkRate) };
                  })
                  .sort((a, b) => b.combined - a.combined);
                if (!ranked.length) return null;
                return (
                  <Card>
                    <div style={{ fontWeight: 700, marginBottom: 8 }}>{tr("Rang objekata po dobiti")}</div>
                    {ranked.map((r, i) => (
                      <div key={r.key} className="num" style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", borderBottom: `1px solid ${S.line}`, fontSize: 13.5 }}>
                        <span style={{ color: S.sub, width: 18 }}>{i + 1}.</span>
                        <span style={{ flex: 1, fontWeight: 600 }}>{r.name}</span>
                        <span style={{ fontWeight: 700, color: r.combined >= 0 ? S.green : S.red }}>
                          {r.profitE !== 0 && eur(r.profitE)}{r.profitE !== 0 && r.profitK !== 0 ? " + " : ""}{r.profitK !== 0 && czk(r.profitK)}
                          {r.profitE === 0 && r.profitK === 0 && eur(0)}
                        </span>
                      </div>
                    ))}
                    <div style={{ fontSize: 11, color: S.sub, marginTop: 6 }}>{tr("Poredano po zbroju u €+Kč (po tečaju iz postavki, samo za rang — knjiži se odvojeno).")}</div>
                  </Card>
                );
              })()}
              {admin && (
            <Card style={{ marginBottom: 12 }}>
              <div onClick={() => setShowPayoutCal(!showPayoutCal)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
                <div style={{ fontWeight: 700 }}>{tr("Pregled isplata (zadnjih 6 mjeseci)")}</div>
                <span style={{ fontWeight: 700, color: S.sub }}>{showPayoutCal ? "▲" : "▼"}</span>
              </div>
              {showPayoutCal && (() => {
                const months = [];
                for (let i = 5; i >= 0; i--) { const d = new Date(y, m - 1 - i, 1); months.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`); }
                const relevantIds = new Set(data.logs.filter((l) => months.includes(monthKey(l.date))).map((l) => l.workerId));
                const ws = sortedWorkers(data.workers.filter((w) => relevantIds.has(w.id)));
                return ws.length === 0 ? (
                  <div style={{ fontSize: 12.5, color: S.sub, marginTop: 8 }}>{tr("Nema radnika sa satima u zadnjih 6 mjeseci.")}</div>
                ) : (
                  <div style={{ overflowX: "auto", marginTop: 10 }}>
                    <table className="num" style={{ borderCollapse: "collapse", width: "100%", fontSize: 12 }}>
                      <thead>
                        <tr>
                          <th style={{ textAlign: "left", padding: "3px 6px", position: "sticky", left: 0, background: S.card }}>{tr("Radnik")}</th>
                          {months.map((mo) => <th key={mo} style={{ padding: "3px 4px", fontWeight: 700 }}>{mShort(Number(mo.slice(5, 7)) - 1)}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {ws.map((w) => (
                          <tr key={w.id}>
                            <td onClick={() => onOpenWorker && onOpenWorker(w.id)} style={{
                              padding: "3px 6px", fontWeight: 600, cursor: onOpenWorker ? "pointer" : "default", color: onOpenWorker ? S.blue : "inherit",
                              position: "sticky", left: 0, background: S.card, whiteSpace: "nowrap" }}>{w.name}</td>
                            {months.map((mo) => {
                              const hadHours = data.logs.some((l) => l.workerId === w.id && monthKey(l.date) === mo);
                              const paid = paidFor(data, w.id, mo);
                              const bg = !hadHours ? "transparent" : paid?.approved ? S.greenSoft : paid ? S.amberSoft : S.redSoft;
                              const color = !hadHours ? S.line : paid?.approved ? S.green : paid ? S.amber : S.red;
                              return (
                                <td key={mo} style={{ textAlign: "center", padding: "3px 2px" }}>
                                  <span style={{ display: "inline-block", width: 20, height: 20, lineHeight: "20px", borderRadius: 5, background: bg, color, fontWeight: 800 }}>
                                    {hadHours ? (paid?.approved ? "✓" : paid ? "⏳" : "✕") : ""}
                                  </span>
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <div style={{ fontSize: 11.5, color: S.sub, marginTop: 8 }}>{tr("✓ zeleno = odobreno i zaključano · ⏳ žuto = predloženo, čeka odobrenje · ✕ crveno = ima sati, ništa predloženo · prazno = nema sati taj mjesec")}</div>
                  </div>
                );
              })()}
            </Card>
              )}
            </>
          )}
        </>
      )}

      {sigTarget && (
        <SignaturePad title={tr("Potpis — {0}", sigTarget.name)}
          onSave={(dataUrl) => { setSignatures((p) => ({ ...p, [sigTarget.id]: dataUrl })); setSigTarget(null); }}
          onClose={() => setSigTarget(null)} />
      )}
    </>
  );
}
