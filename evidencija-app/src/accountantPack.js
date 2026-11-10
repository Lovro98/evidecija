/* ================================================================== */
/*  Paket za knjigovođu: sve izdane i primljene fakture za jedan      */
/*  mjesec (po DUZP-u) — kao ZIP (PDF-ovi + pregled .xlsx) ili kao    */
/*  jedan dokument za ispis (pregled + sve stranice faktura).         */
/*  Bez dodatnih paketa: ZIP se slaže ručno (bez kompresije),         */
/*  PDF stranice se iscrtavaju preko pdf.js koji aplikacija već ima.  */
/* ================================================================== */

import { styledXlsxBuffer, xDate } from "./xlsxStyled.js";

const MONTHS_CZ = ["leden", "únor", "březen", "duben", "květen", "červen", "červenec", "srpen", "září", "říjen", "listopad", "prosinec"];
const monthLabel = (mo) => { const [y, m] = mo.split("-").map(Number); return `${MONTHS_CZ[m - 1]} ${y}`; };
const fmtD = (iso) => { if (!iso) return ""; const [y, m, d] = String(iso).slice(0, 10).split("-"); return `${Number(d)}. ${Number(m)}. ${y}`; };
const fmtMoney = (n, cur) => (Number(n) || 0).toLocaleString("cs-CZ", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + (cur === "EUR" ? " €" : " Kč");
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const safeName = (s) => String(s || "").replace(/[\\/:*?"<>|]+/g, "").replace(/\s+/g, " ").trim().slice(0, 80) || "bez-nazvu";
const extOf = (path) => (String(path).match(/\.([a-z0-9]{2,5})$/i)?.[1] || "pdf").toLowerCase();
const isImage = (path) => /\.(jpe?g|png|gif|webp)$/i.test(path || "");

/* ---------- ZIP (store, UTF-8 imena) ---------- */
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
  return t;
})();
const crc32 = (u8) => { let c = 0xffffffff; for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };

function makeZip(files) {
  const enc = new TextEncoder();
  const now = new Date();
  const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
  const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
  const parts = [], central = [];
  let offset = 0;
  for (const f of files) {
    const name = enc.encode(f.name);
    const data = f.data;
    const crc = crc32(data);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true); local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true);
    local.setUint16(8, 0, true); local.setUint16(10, dosTime, true); local.setUint16(12, dosDate, true);
    local.setUint32(14, crc, true); local.setUint32(18, data.length, true); local.setUint32(22, data.length, true);
    local.setUint16(26, name.length, true); local.setUint16(28, 0, true);
    parts.push(new Uint8Array(local.buffer), name, data);
    const cd = new DataView(new ArrayBuffer(46));
    cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true); cd.setUint16(8, 0x0800, true);
    cd.setUint16(10, 0, true); cd.setUint16(12, dosTime, true); cd.setUint16(14, dosDate, true);
    cd.setUint32(16, crc, true); cd.setUint32(20, data.length, true); cd.setUint32(24, data.length, true);
    cd.setUint16(28, name.length, true); cd.setUint32(42, offset, true);
    central.push(new Uint8Array(cd.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const cdSize = central.reduce((s, p) => s + p.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
  end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
  return new Blob([...parts, ...central, new Uint8Array(end.buffer)], { type: "application/zip" });
}

const download = (blob, fileName) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = fileName;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
};

/* ---------- zajedničko ---------- */
function collect({ out, inn, objName }) {
  const outRows = [...out].sort((a, b) => String(a.number).localeCompare(String(b.number), "cs", { numeric: true })).map((i) => ({
    kind: "out", number: i.number, party: objName(i.objectId), duzp: i.duzp || i.issueDate, issueDate: i.issueDate,
    base: (Number(i.amount) || 0) - (i.vatCurrency === i.currency ? Number(i.vatAmount) || 0 : 0),
    vat: Number(i.vatAmount) || 0, total: Number(i.amount) || 0, currency: i.currency, filePath: i.filePath || "",
  }));
  const inRows = [...inn].sort((a, b) => String(a.duzp).localeCompare(String(b.duzp))).map((p) => ({
    kind: "in", number: p.number, party: p.supplier, duzp: p.duzp, base: Number(p.base) || 0, vat: Number(p.vat) || 0,
    total: round2((Number(p.base) || 0) + (Number(p.vat) || 0)), currency: p.currency, filePath: p.filePath || "", note: p.note || "",
  }));
  return { outRows, inRows };
}
const round2 = (n) => Math.round(n * 100) / 100;

// účtenky (blokovi) za mjesec — iz tablice expenses (oblik mapExpense u App.jsx)
const PAY_CZ = { gotovina: "hotově", kartica: "kartou firmy", banka: "převodem", privatno: "soukromě" };
function collectExp(exp = [], objName) {
  return [...exp].sort((a, b) => String(a.date).localeCompare(String(b.date))).map((e) => ({
    kind: "exp", number: "", party: e.supplier || e.description, description: e.description || "", duzp: e.date,
    base: round2((Number(e.amount) || 0) - (Number(e.vat) || 0)), vat: Number(e.vat) || 0, total: Number(e.amount) || 0,
    currency: e.currency, filePath: e.filePath || "", fileUrl: e.fileUrl || "", pay: PAY_CZ[e.payMethod] || e.payMethod,
    object: e.objectId ? objName(e.objectId) : "", note: [e.paidBy && `platil ${e.paidBy}`, e.note].filter(Boolean).join(" · "),
  }));
}

async function fetchFile(supabase, path) {
  const { data, error } = await supabase.storage.from("invoices").download(path);
  if (error || !data) return null;
  return new Uint8Array(await data.arrayBuffer());
}

/* ---------- 1) ZIP ---------- */
export async function downloadAccountantZip({ supabase, month, out, inn, exp, objName, onProgress }) {
  const { outRows, inRows } = collect({ out, inn, objName });
  const expRows = collectExp(exp, objName);
  const files = [], missing = [], expMissing = [];
  const all = [...outRows, ...inRows, ...expRows];
  let n = 0;
  for (const r of all) {
    onProgress?.(++n, all.length);
    const folder = r.kind === "out" ? "Vydané faktury" : r.kind === "in" ? "Přijaté faktury" : "Účtenky";
    const miss = r.kind === "exp" ? expMissing : missing;
    if (!r.filePath) { miss.push(r); continue; }
    const data = await fetchFile(supabase, r.filePath);
    if (!data) { miss.push(r); continue; }
    const base = r.kind === "out" ? `${safeName(r.number)}_${safeName(r.party)}`
      : r.kind === "exp" ? `${fmtD(r.duzp).replace(/\. /g, "-")}_${safeName(r.party)}_${Math.round(r.total)}`
      : `${fmtD(r.duzp).replace(/\. /g, "-")}_${safeName(r.party)}_${safeName(r.number)}`;
    let name = `${folder}/${base}.${extOf(r.filePath)}`;
    let k = 2; while (files.some((f) => f.name === name)) name = `${folder}/${base}_${k++}.${extOf(r.filePath)}`;
    files.push({ name, data });
  }

  const cols = [
    { key: "number", header: "Číslo faktury", width: 16 },
    { key: "party", header: "", width: 34 },
    { key: "duzp", header: "DUZP", width: 12, type: "date" },
    { key: "base", header: "Základ", width: 14, type: "money" },
    { key: "vat", header: "DPH", width: 13, type: "money" },
    { key: "total", header: "Celkem", width: 14, type: "money" },
    { key: "currency", header: "Měna", width: 7, type: "center" },
    { key: "pdf", header: "PDF v balíčku", width: 13, type: "center" },
  ];
  const toRow = (r) => ({ ...r, duzp: xDate(r.duzp), pdf: missing.includes(r) || expMissing.includes(r) ? (r.fileUrl ? "odkaz" : "CHYBÍ") : "ano" });
  const xbuf = await styledXlsxBuffer({ sheets: [
    { name: "Vydané faktury", title: `Vydané faktury – ${monthLabel(month)}`, columns: cols.map((c) => (c.key === "party" ? { ...c, header: "Odběratel" } : c)),
      rows: outRows.map(toRow), total: { label: "CELKEM", keys: ["base", "vat", "total"] } },
    { name: "Přijaté faktury", title: `Přijaté faktury – ${monthLabel(month)}`, columns: [...cols.map((c) => (c.key === "party" ? { ...c, header: "Dodavatel" } : c)), { key: "note", header: "Poznámka", width: 40, wrap: true }],
      rows: inRows.map(toRow), total: { label: "CELKEM", keys: ["base", "vat", "total"] }, tabColor: "FF2B5D8A" },
    ...(expRows.length ? [{ name: "Účtenky", title: `Účtenky a výdaje – ${monthLabel(month)}`, columns: [
      { key: "duzp", header: "Datum", width: 12, type: "date" }, { key: "party", header: "Dodavatel", width: 28 },
      { key: "description", header: "Popis", width: 26, wrap: true },
      { key: "base", header: "Základ", width: 13, type: "money" }, { key: "vat", header: "DPH", width: 12, type: "money" },
      { key: "total", header: "Celkem", width: 13, type: "money" }, { key: "currency", header: "Měna", width: 7, type: "center" },
      { key: "pay", header: "Placeno", width: 13, type: "center" }, { key: "object", header: "Objekt", width: 16 },
      { key: "pdf", header: "Doklad v balíčku", width: 14, type: "center" }, { key: "note", header: "Poznámka", width: 30, wrap: true },
    ], rows: expRows.map(toRow), total: { label: "CELKEM", keys: ["base", "vat", "total"] }, tabColor: "FFB97F1E" }] : []),
  ] });
  files.unshift({ name: `Přehled faktur ${month}.xlsx`, data: new Uint8Array(xbuf) });

  download(makeZip(files), `Faktury_${month}_Molnar_Rose.zip`);
  return { count: files.length - 1, missing, expMissing };
}

/* ---------- 2) Ispis: pregled + sve stranice faktura ---------- */
export async function printAccountantPack({ supabase, pdfjsLib, month, out, inn, exp, objName, company, onProgress }) {
  const w = window.open("", "_blank");
  if (!w) throw new Error("Preglednik je blokirao novi prozor — dopusti skočne prozore za ovu stranicu.");
  w.document.write(`<!doctype html><meta charset="utf-8"><title>Faktury ${esc(month)}</title><body style="font-family:Arial,sans-serif;padding:24px">Připravuji dokument k tisku…</body>`);

  const { outRows, inRows } = collect({ out, inn, objName });
  const expRows = collectExp(exp, objName);
  const pages = [], missing = [], expMissing = [];
  const all = [...outRows, ...inRows, ...expRows];
  let n = 0;
  for (const r of all) {
    onProgress?.(++n, all.length);
    const miss = r.kind === "exp" ? expMissing : missing;
    if (!r.filePath) { miss.push(r); continue; }
    const data = await fetchFile(supabase, r.filePath);
    if (!data) { miss.push(r); continue; }
    if (isImage(r.filePath)) {
      const url = URL.createObjectURL(new Blob([data]));
      pages.push({ r, src: await new Promise((res) => { const img = new Image(); img.onload = () => { const c = document.createElement("canvas"); c.width = img.naturalWidth; c.height = img.naturalHeight; c.getContext("2d").drawImage(img, 0, 0); res(c.toDataURL("image/jpeg", 0.9)); URL.revokeObjectURL(url); }; img.onerror = () => res(null); img.src = url; }) });
      continue;
    }
    try {
      const pdf = await pdfjsLib.getDocument({ data }).promise;
      for (let p = 1; p <= pdf.numPages; p++) {
        const page = await pdf.getPage(p);
        const vp = page.getViewport({ scale: 2 });
        const c = document.createElement("canvas");
        c.width = vp.width; c.height = vp.height;
        await page.render({ canvasContext: c.getContext("2d"), viewport: vp }).promise;
        pages.push({ r, src: c.toDataURL("image/jpeg", 0.88) });
      }
    } catch { (r.kind === "exp" ? expMissing : missing).push(r); }
  }

  const sum = (rows, k, cur) => rows.filter((r) => r.currency === cur).reduce((s, r) => s + r[k], 0);
  const totals = (rows) => ["CZK", "EUR"].filter((c) => rows.some((r) => r.currency === c))
    .map((c) => `<tr class="tot"><td colspan="3">CELKEM ${c}</td><td>${fmtMoney(sum(rows, "base", c), c)}</td><td>${fmtMoney(sum(rows, "vat", c), c)}</td><td>${fmtMoney(sum(rows, "total", c), c)}</td><td></td></tr>`).join("");
  const table = (title, partyLabel, rows) => `
    <h2>${title} (${rows.length})</h2>
    ${rows.length === 0 ? `<p class="sub">Žádné faktury.</p>` : `<table>
      <tr><th>Číslo</th><th>${partyLabel}</th><th>DUZP</th><th>Základ</th><th>DPH</th><th>Celkem</th><th>PDF</th></tr>
      ${rows.map((r) => `<tr><td>${esc(r.kind === "exp" ? r.pay : r.number)}</td><td>${esc(r.party)}${r.kind === "exp" && r.description && r.description !== r.party ? ` <span class="sub">${esc(r.description)}</span>` : ""}</td><td>${fmtD(r.duzp)}</td><td>${fmtMoney(r.base, r.currency)}</td><td>${fmtMoney(r.vat, r.currency)}</td><td>${fmtMoney(r.total, r.currency)}</td><td>${missing.includes(r) || expMissing.includes(r) ? '<b class="miss">chybí</b>' : "✓"}</td></tr>`).join("")}
      ${totals(rows)}</table>`}`;

  const html = `<!doctype html><html><head><meta charset="utf-8"><title>Faktury ${esc(monthLabel(month))} – ${esc(company)}</title>
<style>
  @page { size: A4; margin: 12mm; }
  body { font-family: Arial, sans-serif; color: #1c2521; margin: 0; }
  .cover { padding: 8mm 4mm; }
  h1 { font-size: 20px; margin: 0 0 4px; } h2 { font-size: 15px; margin: 18px 0 6px; color: #0e6b4f; }
  .sub { color: #5c6862; font-size: 12px; }
  table { width: 100%; border-collapse: collapse; font-size: 11.5px; }
  th, td { border-bottom: 1px solid #d5dbd3; padding: 5px 4px; text-align: left; }
  td:nth-child(n+4), th:nth-child(n+4) { text-align: right; white-space: nowrap; }
  th { background: #e3f0ea; } .tot td { font-weight: 700; border-top: 2px solid #0e6b4f; }
  .miss { color: #b3402f; }
  .page { page-break-before: always; text-align: center; }
  .page img { max-width: 100%; max-height: 270mm; }
  .label { font-size: 10px; color: #5c6862; text-align: left; margin-bottom: 2mm; }
  .bar { position: sticky; top: 0; background: #fff; padding: 10px; border-bottom: 1px solid #d5dbd3; text-align: right; }
  .bar button { font-size: 15px; padding: 8px 18px; border-radius: 10px; border: 0; background: #0e6b4f; color: #fff; cursor: pointer; }
  @media print { .bar { display: none; } }
</style></head><body>
<div class="bar"><button onclick="window.print()">🖨️ Tisk / uložit jako PDF</button></div>
<div class="cover">
  <h1>Faktury za ${esc(monthLabel(month))}</h1>
  <div class="sub">${esc(company)} · podle DUZP · vytvořeno ${fmtD(new Date().toISOString().slice(0, 10))}</div>
  ${table("Vydané faktury", "Odběratel", outRows)}
  ${table("Přijaté faktury", "Dodavatel", inRows)}
  ${expRows.length ? table("Účtenky a výdaje (sloupec Číslo = způsob platby)", "Dodavatel", expRows) : ""}
  ${missing.length ? `<p class="miss" style="margin-top:14px">⚠️ U ${missing.length} faktur chybí PDF v Evidenci — nejsou vytištěny: ${missing.map((r) => esc(r.number || r.party)).join(", ")}.</p>` : ""}
</div>
${pages.filter((p) => p.src).map((p) => `<div class="page"><div class="label">${p.r.kind === "out" ? "Vydaná faktura" : p.r.kind === "in" ? "Přijatá faktura" : "Účtenka"} ${esc(p.r.number)} · ${esc(p.r.party)}</div><img src="${p.src}"></div>`).join("")}
</body></html>`;
  w.document.open(); w.document.write(html); w.document.close();
  return { pages: pages.length, missing, expMissing };
}
