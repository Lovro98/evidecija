/* ================================================================== */
/*  Zajednički izgled svih Excel izvoza (ExcelJS)                      */
/*  naslov · zaglavlje u boji · formati valuta · zbrojevi formulama   */
/*  filteri · zamrznuto zaglavlje · postavke ispisa                    */
/* ================================================================== */

export const XC = {
  green: "FF0E6B4F", greenSoft: "FFE3F0EA", ink: "FF1C2521", sub: "FF5C6862", line: "FFD5DBD3",
  band: "FFF7F9F6", amber: "FFB97F1E", amberSoft: "FFFBF1DC", red: "FFB3402F", redSoft: "FFF9E8E4",
  blue: "FF2B5D8A", white: "FFFFFFFF", muted: "FF9AA39E",
};
const FONT = "Arial";
const FMT = {
  eur: '#,##0.00 "€";-#,##0.00 "€";"–"',
  czk: '#,##0 "Kč";-#,##0 "Kč";"–"',
  money: '#,##0.00;-#,##0.00;"–"',
  hours: '#,##0.00;-#,##0.00;"–"',
  int: '#,##0;-#,##0;"–"',
  date: "dd.mm.yyyy",
};
const RIGHT = new Set(["eur", "czk", "money", "hours", "int"]);
const thin = { style: "thin", color: { argb: XC.line } };
const box = { top: thin, left: thin, bottom: thin, right: thin };
const colLetter = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const todayLabel = () => { const d = new Date(); return `${d.getDate()}. ${d.getMonth() + 1}. ${d.getFullYear()}.`; };

/** ISO "2026-05-17" → Date (UTC) za prave datume u Excelu */
export const xDate = (iso) => { if (!iso) return null; const [y, m, d] = String(iso).slice(0, 10).split("-").map(Number); if (!y || !m || !d) return null; return new Date(Date.UTC(y, m - 1, d)); };

/**
 * Jedan list.
 * columns: [{ key, header, width, type: text|eur|czk|money|hours|int|date|center, wrap, bold, color(v,row)→argb, fill(v,row)→argb }]
 * rows:    [{ [key]: value }]
 * total:   { label, keys: [...] } → red UKUPNO sa SUM formulama
 */
function addSheet(wb, sh) {
  const cols = sh.columns;
  const ws = wb.addWorksheet(sh.name.slice(0, 31), {
    properties: { tabColor: { argb: sh.tabColor || XC.green } },
    pageSetup: { paperSize: 9, orientation: sh.landscape ? "landscape" : "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 } },
  });
  ws.columns = cols.map((c) => ({ key: c.key, width: c.width || 14 }));

  // naslov
  let headerRow = 1;
  if (!sh.noTitle) {
    ws.mergeCells(1, 1, 1, Math.max(cols.length, 3));
    ws.mergeCells(2, 1, 2, Math.max(cols.length, 3));
    const t = ws.getCell(1, 1);
    t.value = sh.title || sh.name;
    t.font = { name: FONT, bold: true, size: 15, color: { argb: XC.green } };
    ws.getRow(1).height = 26;
    const s = ws.getCell(2, 1);
    s.value = [sh.subtitle, `izrađeno ${todayLabel()}`].filter(Boolean).join("  ·  ");
    s.font = { name: FONT, size: 10, italic: true, color: { argb: XC.sub } };
    ws.addRow([]);
    headerRow = 4;
  }

  // zaglavlje
  const h = ws.addRow(cols.map((c) => c.header));
  h.height = 30;
  h.eachCell((c) => {
    c.font = { name: FONT, bold: true, size: 10, color: { argb: XC.white } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: sh.headerColor || XC.green } };
    c.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    c.border = box;
  });

  // podaci
  const first = headerRow + 1;
  const rows = sh.rows || [];
  rows.forEach((r, i) => {
    const row = ws.addRow(cols.map((c) => {
      const v = r[c.key];
      if (c.type === "date") return v instanceof Date ? v : (xDate(v) || (v ?? ""));
      return v ?? "";
    }));
    row.eachCell({ includeEmpty: true }, (cell, ci) => {
      const c = cols[ci - 1];
      if (!c) return;
      const v = r[c.key];
      cell.font = { name: FONT, size: 10, bold: !!c.bold, color: { argb: (c.color && c.color(v, r)) || XC.ink } };
      cell.border = box;
      const fill = (c.fill && c.fill(v, r)) || (i % 2 === 1 ? XC.band : null);
      if (fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
      if (FMT[c.type]) cell.numFmt = FMT[c.type];
      cell.alignment = { vertical: "middle", wrapText: !!c.wrap,
        horizontal: RIGHT.has(c.type) && typeof v === "number" ? "right" : (c.type === "date" || c.type === "center") ? "center" : "left" };
    });
  });
  if (!rows.length) {
    const row = ws.addRow([sh.emptyText || "Nema podataka."]);
    row.getCell(1).font = { name: FONT, italic: true, color: { argb: XC.sub } };
  }
  const last = Math.max(first, first + rows.length - 1);

  // UKUPNO
  if (sh.total && rows.length) {
    const vals = cols.map((c, ci) => {
      if (ci === 0) return sh.total.label || "UKUPNO";
      if (!sh.total.keys.includes(c.key)) return "";
      const L = colLetter(ci + 1);
      const result = rows.reduce((s, r) => s + (typeof r[c.key] === "number" ? r[c.key] : 0), 0);
      return { formula: `SUBTOTAL(9,${L}${first}:${L}${last})`, result };
    });
    const t = ws.addRow(vals);
    t.height = 20;
    t.eachCell({ includeEmpty: true }, (cell, ci) => {
      const c = cols[ci - 1];
      cell.font = { name: FONT, bold: true, size: 10, color: { argb: XC.ink } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: XC.greenSoft } };
      cell.border = { top: { style: "medium", color: { argb: XC.green } }, bottom: thin, left: thin, right: thin };
      if (c && FMT[c.type]) { cell.numFmt = FMT[c.type]; cell.alignment = { horizontal: "right" }; }
    });
  }

  // napomena ispod tablice
  if (sh.note) {
    ws.addRow([]);
    const n = ws.addRow([sh.note]);
    ws.mergeCells(n.number, 1, n.number, Math.max(cols.length, 3));
    n.getCell(1).font = { name: FONT, size: 8.5, italic: true, color: { argb: XC.sub } };
    n.getCell(1).alignment = { wrapText: true, vertical: "top" };
    n.height = 30;
  }

  ws.views = [{ state: "frozen", ySplit: headerRow, xSplit: sh.freezeCols || 0 }];
  if (rows.length) ws.autoFilter = { from: { row: headerRow, column: 1 }, to: { row: last, column: cols.length } };
  ws.pageSetup.printTitlesRow = `${headerRow}:${headerRow}`;
  ws.headerFooter = { oddFooter: `&L&8Molnar Rose agency & floristry s.r.o.&C&8${(sh.title || sh.name).replace(/&/g, "&&")}&R&8Stranica &P / &N` };
  if (sh.after) sh.after(ws, { headerRow, first, last });
  return ws;
}

/** Izradi i preuzmi .xlsx s jednim ili više listova. */
export async function downloadStyledXlsx({ fileName, sheets }) {
  const { default: ExcelJS } = await import("exceljs");
  const wb = new ExcelJS.Workbook();
  wb.creator = "Evidencija rada – Molnar Rose";
  wb.created = new Date();
  wb.calcProperties.fullCalcOnLoad = true;
  sheets.forEach((sh) => addSheet(wb, sh));
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = fileName;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
