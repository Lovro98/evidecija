/* ================================================================== */
/*  NEISPLAĆENO — izvoz u Excel s formatiranjem (ExcelJS)              */
/*  Listovi: Pregled (zbrojevi formulama), Koverte, Isplate            */
/* ================================================================== */

const C = {
  green: "FF0E6B4F", greenSoft: "FFE3F0EA", ink: "FF1C2521", sub: "FF5C6862",
  line: "FFD5DBD3", band: "FFF7F9F6", amber: "FFB97F1E", amberSoft: "FFFBF1DC", white: "FFFFFFFF",
};
const KC = '#,##0 "Kč";-#,##0 "Kč";"–"';
const FONT = "Arial";

const monthLabel = (mo) => { if (!mo) return ""; const [y, m] = mo.split("-"); return `${m}/${y}`; };
const toDate = (iso) => { if (!iso) return null; const [y, m, d] = iso.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const fmtToday = () => { const d = new Date(); return `${d.getDate()}. ${d.getMonth() + 1}. ${d.getFullYear()}.`; };
const thin = { style: "thin", color: { argb: C.line } };
const allBorders = { top: thin, left: thin, bottom: thin, right: thin };

function styleHeader(row) {
  row.height = 22;
  row.eachCell((c) => {
    c.font = { name: FONT, bold: true, size: 10, color: { argb: C.white } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: C.green } };
    c.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    c.border = allBorders;
  });
}
function styleBody(row, i, moneyCols, opts = {}) {
  row.eachCell({ includeEmpty: true }, (c, col) => {
    c.font = { name: FONT, size: 10, color: { argb: C.ink } };
    c.border = allBorders;
    c.alignment = { vertical: "middle", wrapText: !!opts.wrapCols?.includes(col) };
    if (i % 2 === 1) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: C.band } };
    if (moneyCols.includes(col)) { c.numFmt = KC; c.alignment = { vertical: "middle", horizontal: "right" }; }
  });
}
function styleTotal(row, moneyCols) {
  row.height = 20;
  row.eachCell({ includeEmpty: true }, (c, col) => {
    c.font = { name: FONT, bold: true, size: 10, color: { argb: C.ink } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: C.greenSoft } };
    c.border = { top: { style: "medium", color: { argb: C.green } }, bottom: thin, left: thin, right: thin };
    if (moneyCols.includes(col)) { c.numFmt = KC; c.alignment = { horizontal: "right" }; }
  });
}
function titleBlock(ws, title, subtitle, lastCol) {
  ws.mergeCells(1, 1, 1, lastCol);
  ws.mergeCells(2, 1, 2, lastCol);
  const t = ws.getCell(1, 1);
  t.value = title;
  t.font = { name: FONT, bold: true, size: 15, color: { argb: C.green } };
  ws.getRow(1).height = 26;
  const s = ws.getCell(2, 1);
  s.value = subtitle;
  s.font = { name: FONT, size: 10, italic: true, color: { argb: C.sub } };
  ws.addRow([]);
}
function pageSetup(ws, landscape, headerRow) {
  ws.pageSetup = { paperSize: 9, orientation: landscape ? "landscape" : "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 }, printTitlesRow: `${headerRow}:${headerRow}` };
  ws.headerFooter = { oddFooter: "&L&8Molnar Rose agency & floristry s.r.o.&R&8Stranica &P / &N" };
}

/**
 * rows:  [{ id, worker_name, objectName, month, section, amount, paid, remaining, calc_note, note }]
 * pays:  [{ unpaid_wage_id, pay_date, amount, note }]
 * scope: naziv objekta ili null za sve objekte
 */
export async function exportUnpaidXlsx({ rows, pays, scope, fileName }) {
  const { default: ExcelJS } = await import("exceljs");
  const wb = new ExcelJS.Workbook();
  wb.creator = "Evidencija rada – Molnar Rose";
  wb.created = new Date();
  wb.calcProperties.fullCalcOnLoad = true;

  const sorted = [...rows].sort((a, b) => a.objectName.localeCompare(b.objectName, "hr") || a.worker_name.localeCompare(b.worker_name, "hr") || a.month.localeCompare(b.month));
  const byId = Object.fromEntries(sorted.map((r) => [r.id, r]));
  const payRows = pays.filter((p) => byId[p.unpaid_wage_id])
    .sort((a, b) => (a.pay_date || "").localeCompare(b.pay_date || "") || byId[a.unpaid_wage_id].worker_name.localeCompare(byId[b.unpaid_wage_id].worker_name, "hr"));
  const scopeText = scope ? `Objekt: ${scope}` : "Svi objekti";
  const subtitle = `${scopeText}  ·  stanje na dan ${fmtToday()}  ·  iznosi u Kč`;

  const wsP = wb.addWorksheet("Pregled", { properties: { tabColor: { argb: C.green } }, views: [{ showGridLines: false }] });
  const wsK = wb.addWorksheet("Koverte", { properties: { tabColor: { argb: C.amber } } });
  const wsI = wb.addWorksheet("Isplate", { properties: { tabColor: { argb: C.sub } } });

  /* ---------- Koverte ---------- */
  const kCols = [
    { key: "obj", header: "Objekt", width: 18 }, { key: "w", header: "Radnik", width: 26 }, { key: "m", header: "Mjesec", width: 10 },
    { key: "sec", header: "Odjel", width: 16 }, { key: "amt", header: "Iznos", width: 13 }, { key: "paid", header: "Isplaćeno", width: 13 },
    { key: "rem", header: "Ostaje", width: 13 }, { key: "calc", header: "Izračun s koverte", width: 58 }, { key: "note", header: "Napomena", width: 34 },
    { key: "id", header: "ID koverte", width: 12 },
  ];
  wsK.columns = kCols.map((c) => ({ key: c.key, width: c.width }));
  titleBlock(wsK, "Neisplaćene plaće – koverte", subtitle, kCols.length);
  const kHead = wsK.addRow(kCols.map((c) => c.header)); styleHeader(kHead);
  const kHeadRow = kHead.number;
  const kFirst = kHeadRow + 1;
  const kLast = kFirst + sorted.length - 1;
  const iFirst = 5;
  const iLast = Math.max(iFirst, iFirst + payRows.length - 1);
  sorted.forEach((r, i) => {
    const n = kFirst + i;
    const row = wsK.addRow([
      r.objectName, r.worker_name, monthLabel(r.month), r.section || "", r.amount,
      { formula: `SUMIFS(Isplate!$E$${iFirst}:$E$${iLast},Isplate!$G$${iFirst}:$G$${iLast},$J${n})`, result: r.paid },
      { formula: `E${n}-F${n}`, result: r.remaining },
      r.calc_note || "", r.note || "", r.id,
    ]);
    styleBody(row, i, [5, 6, 7], { wrapCols: [8, 9] });
    row.getCell(3).alignment = { horizontal: "center", vertical: "middle" };
    row.getCell(10).font = { name: FONT, size: 8, color: { argb: "FF9AA39E" } };
    row.getCell(7).font = { name: FONT, size: 10, bold: true, color: { argb: r.remaining > 0.004 ? C.amber : C.green } };
  });
  const kTot = wsK.addRow(["UKUPNO", "", "", "",
    { formula: `SUM(E${kFirst}:E${kLast})`, result: sorted.reduce((s, r) => s + r.amount, 0) },
    { formula: `SUM(F${kFirst}:F${kLast})`, result: sorted.reduce((s, r) => s + r.paid, 0) },
    { formula: `SUM(G${kFirst}:G${kLast})`, result: sorted.reduce((s, r) => s + r.remaining, 0) }, "", "", ""]);
  styleTotal(kTot, [5, 6, 7]);
  wsK.views = [{ state: "frozen", ySplit: kHeadRow, xSplit: 2 }];
  wsK.autoFilter = { from: { row: kHeadRow, column: 1 }, to: { row: kLast, column: kCols.length } };
  wsK.getColumn(10).outlineLevel = 1;
  pageSetup(wsK, true, kHeadRow);

  /* ---------- Isplate ---------- */
  const iCols = [
    { key: "d", header: "Datum", width: 13 }, { key: "obj", header: "Objekt", width: 18 }, { key: "w", header: "Radnik", width: 26 },
    { key: "m", header: "Za mjesec", width: 11 }, { key: "amt", header: "Iznos", width: 13 }, { key: "note", header: "Napomena", width: 44 },
    { key: "id", header: "ID koverte", width: 12 },
  ];
  wsI.columns = iCols.map((c) => ({ key: c.key, width: c.width }));
  titleBlock(wsI, "Isplate neisplaćenih plaća", subtitle, iCols.length);
  const iHead = wsI.addRow(iCols.map((c) => c.header)); styleHeader(iHead);
  payRows.forEach((p, i) => {
    const r = byId[p.unpaid_wage_id];
    const row = wsI.addRow([toDate(p.pay_date) || "bez datuma", r.objectName, r.worker_name, monthLabel(r.month), p.amount, p.note || "", r.id]);
    styleBody(row, i, [5], { wrapCols: [6] });
    row.getCell(1).numFmt = "dd.mm.yyyy";
    row.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
    row.getCell(4).alignment = { horizontal: "center", vertical: "middle" };
    row.getCell(7).font = { name: FONT, size: 8, color: { argb: "FF9AA39E" } };
  });
  if (!payRows.length) { const row = wsI.addRow(["Još nema isplata."]); row.getCell(1).font = { name: FONT, italic: true, color: { argb: C.sub } }; }
  const iTot = wsI.addRow(["UKUPNO", "", "", "", { formula: `SUM(E${iFirst}:E${iLast})`, result: payRows.reduce((s, p) => s + p.amount, 0) }, "", ""]);
  styleTotal(iTot, [5]);
  wsI.views = [{ state: "frozen", ySplit: 4 }];
  if (payRows.length) wsI.autoFilter = { from: { row: 4, column: 1 }, to: { row: iLast, column: iCols.length } };
  wsI.getColumn(7).outlineLevel = 1;
  pageSetup(wsI, false, 4);

  /* ---------- Pregled ---------- */
  wsP.columns = [{ width: 24 }, { width: 28 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 3 }];
  titleBlock(wsP, "Neisplaćene plaće – Molnar Rose", subtitle, 5);
  const K = (col) => `Koverte!$${col}$${kFirst}:$${col}$${kLast}`;
  const totAmt = sorted.reduce((s, r) => s + r.amount, 0);
  const totPaid = sorted.reduce((s, r) => s + r.paid, 0);
  const kpi = [
    ["Na kovertama ukupno", { formula: `SUM(${K("E")})`, result: totAmt }, false],
    ["Već isplaćeno", { formula: `SUM(${K("F")})`, result: totPaid }, false],
    ["Ukupno dugujem", { formula: `SUM(${K("G")})`, result: totAmt - totPaid }, true],
  ];
  kpi.forEach(([label, val, strong]) => {
    const row = wsP.addRow([label, val]);
    row.height = strong ? 24 : 18;
    const a = row.getCell(1), b = row.getCell(2);
    a.font = { name: FONT, size: strong ? 12 : 10, bold: strong, color: { argb: strong ? C.ink : C.sub } };
    b.font = { name: FONT, size: strong ? 14 : 11, bold: true, color: { argb: strong ? C.amber : C.ink } };
    b.numFmt = KC; b.alignment = { horizontal: "left", vertical: "middle" }; a.alignment = { vertical: "middle" };
    if (strong) [a, b].forEach((c) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: C.amberSoft } }; c.border = { top: { style: "thin", color: { argb: "FFEBD9AE" } }, bottom: { style: "thin", color: { argb: "FFEBD9AE" } } }; });
  });
  wsP.addRow([]);

  const section = (label) => {
    const r = wsP.addRow([label]);
    r.getCell(1).font = { name: FONT, bold: true, size: 12, color: { argb: C.green } };
    r.height = 20;
  };

  // po objektima (samo kad su svi objekti)
  const objects = [...new Set(sorted.map((r) => r.objectName))];
  if (!scope && objects.length) {
    section("Po objektima");
    const h = wsP.addRow(["Objekt", "Radnika s dugom", "Na kovertama", "Isplaćeno", "Ostaje"]); styleHeader(h);
    const first = h.number + 1;
    objects.sort((a, b) => {
      const ra = sorted.filter((r) => r.objectName === a).reduce((s, r) => s + r.remaining, 0);
      const rb = sorted.filter((r) => r.objectName === b).reduce((s, r) => s + r.remaining, 0);
      return rb - ra;
    }).forEach((o, i) => {
      const n = first + i;
      const list = sorted.filter((r) => r.objectName === o);
      const openW = new Set(list.filter((r) => r.remaining > 0.004).map((r) => r.worker_name)).size;
      const row = wsP.addRow([o, openW,
        { formula: `SUMIFS(${K("E")},${K("A")},$A${n})`, result: list.reduce((s, r) => s + r.amount, 0) },
        { formula: `SUMIFS(${K("F")},${K("A")},$A${n})`, result: list.reduce((s, r) => s + r.paid, 0) },
        { formula: `C${n}-D${n}`, result: list.reduce((s, r) => s + r.remaining, 0) }]);
      styleBody(row, i, [3, 4, 5]);
      row.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      row.getCell(5).font = { name: FONT, size: 10, bold: true, color: { argb: C.amber } };
    });
    const last = first + objects.length - 1;
    const t = wsP.addRow(["UKUPNO", "", { formula: `SUM(C${first}:C${last})`, result: totAmt }, { formula: `SUM(D${first}:D${last})`, result: totPaid }, { formula: `SUM(E${first}:E${last})`, result: totAmt - totPaid }]);
    styleTotal(t, [3, 4, 5]);
    wsP.addRow([]);
  }

  // po radnicima
  section("Po radnicima");
  const h2 = wsP.addRow(["Objekt", "Radnik", "Na kovertama", "Isplaćeno", "Ostaje"]); styleHeader(h2);
  const wFirst = h2.number + 1;
  const pairs = [];
  sorted.forEach((r) => { if (!pairs.find((p) => p.o === r.objectName && p.w === r.worker_name)) pairs.push({ o: r.objectName, w: r.worker_name }); });
  pairs.forEach((p, i) => {
    const n = wFirst + i;
    const list = sorted.filter((r) => r.objectName === p.o && r.worker_name === p.w);
    const row = wsP.addRow([p.o, p.w,
      { formula: `SUMIFS(${K("E")},${K("A")},$A${n},${K("B")},$B${n})`, result: list.reduce((s, r) => s + r.amount, 0) },
      { formula: `SUMIFS(${K("F")},${K("A")},$A${n},${K("B")},$B${n})`, result: list.reduce((s, r) => s + r.paid, 0) },
      { formula: `C${n}-D${n}`, result: list.reduce((s, r) => s + r.remaining, 0) }]);
    styleBody(row, i, [3, 4, 5]);
    const rem = list.reduce((s, r) => s + r.remaining, 0);
    row.getCell(5).font = { name: FONT, size: 10, bold: true, color: { argb: rem > 0.004 ? C.amber : C.green } };
  });
  const wLast = wFirst + pairs.length - 1;
  const t2 = wsP.addRow(["UKUPNO", "", { formula: `SUM(C${wFirst}:C${wLast})`, result: totAmt }, { formula: `SUM(D${wFirst}:D${wLast})`, result: totPaid }, { formula: `SUM(E${wFirst}:E${wLast})`, result: totAmt - totPaid }]);
  styleTotal(t2, [3, 4, 5]);
  wsP.addRow([]);
  const note = wsP.addRow(["Zbrojevi su formule nad listom „Koverte”, a „Isplaćeno” se računa iz lista „Isplate”. Isplate se upisuju u aplikaciji Evidencija rada."]);
  wsP.mergeCells(note.number, 1, note.number, 5);
  note.getCell(1).font = { name: FONT, size: 8, italic: true, color: { argb: C.sub } };
  note.getCell(1).alignment = { wrapText: true };
  note.height = 26;
  pageSetup(wsP, false, h2.number);

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = fileName;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
