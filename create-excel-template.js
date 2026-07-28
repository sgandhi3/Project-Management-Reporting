// Creates template.xlsx — the starter Excel template for the MMO Report Generator.
//
// Run once:  node create-excel-template.js
//
// After generation, open template.xlsx in Excel to:
//   • Rearrange / reformat any sheet
//   • Add charts (recommended: reference 'By Workstream'!A1:H5, 'Sub-Suites'!A1:H30, etc.)
//   • Add pivot tables, slicers, or extra sheets
//
// The excel extension (extensions/excel.js) will load this file each run,
// replace {{TOKEN}} placeholders with live data, and write the populated output.

import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUTPUT = path.join(__dirname, 'template.xlsx');

// ─── Colors ───────────────────────────────────────────────────────────────────
const ACCENT      = 'FF2563EB';
const WHITE       = 'FFFFFFFF';
const GREEN_FILL  = 'FFD1FAE5';
const RED_FILL    = 'FFFEE2E2';
const YELLOW_FILL = 'FFFEF3C7';
const ALT_ROW     = 'FFF8FAFC';
const HEADER_FILL = 'FFE0E7FF';
const MUTED       = 'FF6B7280';

// ─── Style helpers ────────────────────────────────────────────────────────────

function font(bold = false, size = 11, color = null) {
  return { name: 'Calibri', size, bold, ...(color ? { color: { argb: color } } : {}) };
}

function fill(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function border() {
  const thin = { style: 'thin', color: { argb: 'FFD1D5DB' } };
  return { top: thin, left: thin, bottom: thin, right: thin };
}

function freeze(sheet, ySplit = 1) {
  sheet.views = [{ state: 'frozen', ySplit }];
}

function widths(sheet, ws) {
  ws.forEach((w, i) => { sheet.getColumn(i + 1).width = w; });
}

function headerCell(cell, text) {
  cell.value = text;
  cell.font  = font(true, 12, WHITE);
  cell.fill  = fill(ACCENT);
  cell.border = border();
  cell.alignment = { vertical: 'middle', horizontal: 'center' };
}

function dataCell(cell, value, { bold = false, fillArgb = null, align = 'left' } = {}) {
  cell.value = value;
  cell.font  = font(bold, 11);
  cell.border = border();
  cell.alignment = { vertical: 'middle', horizontal: align };
  if (fillArgb) cell.fill = fill(fillArgb);
}

// ─── Sheet 1: Summary ─────────────────────────────────────────────────────────

function buildSummary(wb) {
  const ws = wb.addWorksheet('Summary');
  widths(ws, [28, 22, 16, 16, 16, 16]);
  freeze(ws);

  let r = 1;

  // Title
  ws.getRow(r).height = 30;
  const title = ws.getCell(`A${r}`);
  title.value = 'Weekly SIT Status Report';
  title.font  = font(true, 18, WHITE);
  title.fill  = fill(ACCENT);
  title.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.mergeCells(`A${r}:F${r}`);
  r++;

  // Metadata row
  ws.getRow(r).height = 18;
  const meta = [
    ['Report Date',       '{{Date}}'],
    ['Reporting Period',  '{{PP}}'],
    ['Forecast Period',   '{{FP}}'],
  ];
  meta.forEach(([label, token], i) => {
    const col = i * 2 + 1; // A, C, E
    const lc  = ws.getCell(r, col);
    const vc  = ws.getCell(r, col + 1);
    lc.value = label + ':';
    lc.font  = font(true, 11);
    vc.value = token;
    vc.font  = font(false, 11);
    vc.fill  = fill(ALT_ROW);
    vc.border = border();
  });
  r += 2;

  // KEY METRICS section
  const km = ws.getCell(`A${r}`);
  km.value = 'KEY METRICS';
  km.font  = font(true, 13, ACCENT.slice(2));
  r++;

  // Metric table header
  ['Metric', 'Value'].forEach((h, i) => headerCell(ws.getCell(r, i + 1), h));
  r++;

  const metrics = [
    ['Total Planned',   '{{TotalPlanned}}',   null],
    ['Total Executed',  '{{TotalExecuted}}',  null],
    ['Total Passed',    '{{TotalPassed}}',    GREEN_FILL],
    ['Total Failed',    '{{TotalFailed}}',    RED_FILL],
    ['Pass Rate',       '{{PassRate}}',       null],
    ['Not Started',     '{{TotalNotStarted}}',null],
    ['In Progress',     '{{TotalInProgress}}',null],
  ];

  metrics.forEach(([label, token, rowFill], idx) => {
    const argb = rowFill || (idx % 2 === 0 ? ALT_ROW : null);
    dataCell(ws.getCell(r, 1), label, { fillArgb: argb });
    dataCell(ws.getCell(r, 2), token, { bold: true, fillArgb: argb, align: 'center' });
    r++;
  });

  r++; // blank

  // AI Summary section
  const aiLabel = ws.getCell(`A${r}`);
  aiLabel.value = 'AI SUMMARY';
  aiLabel.font  = font(true, 13, ACCENT.slice(2));
  r++;

  const aiCell = ws.getCell(`A${r}`);
  aiCell.value = '{{AISummary}}';
  aiCell.font  = font(false, 11);
  aiCell.fill  = fill(HEADER_FILL);
  aiCell.border = border();
  aiCell.alignment = { wrapText: true, vertical: 'top' };
  ws.mergeCells(`A${r}:F${r}`);
  ws.getRow(r).height = 80;
  r += 2;

  // Notes
  const note = ws.getCell(`A${r}`);
  note.value = 'Tokens in {{ }} are replaced automatically. Add charts below referencing the "By Workstream" and "Sub-Suites" sheets.';
  note.font  = { name: 'Calibri', size: 10, italic: true, color: { argb: MUTED } };
  ws.mergeCells(`A${r}:F${r}`);
}

// ─── Sheet 2: By Workstream ───────────────────────────────────────────────────

function buildByWorkstream(wb) {
  const ws = wb.addWorksheet('By Workstream');
  widths(ws, [22, 12, 12, 12, 12, 14, 14, 12]);
  freeze(ws);

  const headers = ['Workstream', 'Planned', 'Executed', 'Passed', 'Failed', 'Not Started', 'In Progress', 'Pass %'];
  headers.forEach((h, i) => headerCell(ws.getCell(1, i + 1), h));
  ws.getRow(1).height = 22;
  ws.autoFilter = { from: 'A1', to: 'H1' };

  // Demo data rows (replaced by extension at run time)
  const demo = [
    ['PDM',        120, 95, 80, 15, 25, 0],
    ['Benefits',    85, 70, 65,  5, 15, 0],
    ['Enrollment', 100, 88, 75, 13, 12, 0],
    ['EDI',         60, 45, 40,  5, 15, 0],
  ];

  demo.forEach(([ws_name, pl, ex, pa, fa, ns, ip], idx) => {
    const row = ws.getRow(idx + 2);
    const passP = ex ? Math.round((pa / ex) * 100) : 0;
    const rowFill = idx % 2 === 0 ? ALT_ROW : null;

    dataCell(row.getCell(1), ws_name, { fillArgb: rowFill });
    dataCell(row.getCell(2), pl,      { fillArgb: rowFill, align: 'center' });
    dataCell(row.getCell(3), ex,      { fillArgb: rowFill, align: 'center' });
    dataCell(row.getCell(4), pa,      { fillArgb: pa > 0 ? GREEN_FILL : rowFill, align: 'center' });
    dataCell(row.getCell(5), fa,      { fillArgb: fa > 0 ? RED_FILL   : rowFill, align: 'center' });
    dataCell(row.getCell(6), ns,      { fillArgb: rowFill, align: 'center' });
    dataCell(row.getCell(7), ip,      { fillArgb: rowFill, align: 'center' });
    const passFill = passP >= 80 ? GREEN_FILL : passP >= 60 ? YELLOW_FILL : RED_FILL;
    dataCell(row.getCell(8), `${passP}%`, { fillArgb: passFill, align: 'center' });
  });

  // Total row
  const totalRow = ws.getRow(6);
  const totals   = ['TOTAL', 365, 298, 260, 38, 67, 0];
  totals.forEach((v, i) => {
    const cell = totalRow.getCell(i + 1);
    cell.value  = i === 0 ? v : v;
    cell.font   = font(true, 11);
    cell.fill   = fill(HEADER_FILL);
    cell.border = border();
    cell.alignment = { horizontal: i === 0 ? 'left' : 'center' };
  });
  const totalPassCell = totalRow.getCell(8);
  totalPassCell.value = '87%';
  totalPassCell.font  = font(true, 11);
  totalPassCell.fill  = fill(GREEN_FILL);
  totalPassCell.border = border();
  totalPassCell.alignment = { horizontal: 'center' };

  // Guide note below
  ws.getRow(8).getCell(1).value = '↑ Add charts referencing rows 2:5 of this sheet. The extension overwrites this data on each run.';
  ws.getRow(8).getCell(1).font  = { name: 'Calibri', size: 10, italic: true, color: { argb: MUTED } };
  ws.mergeCells('A8:H8');
}

// ─── Sheet 3: Sub-Suites ──────────────────────────────────────────────────────

function buildSubSuites(wb) {
  const ws = wb.addWorksheet('Sub-Suites');
  widths(ws, [40, 12, 12, 12, 12, 14, 14, 12]);
  freeze(ws);

  const headers = ['Suite', 'Planned', 'Executed', 'Passed', 'Failed', 'Not Started', 'In Progress', 'Pass %'];
  headers.forEach((h, i) => headerCell(ws.getCell(1, i + 1), h));
  ws.getRow(1).height = 22;
  ws.autoFilter = { from: 'A1', to: 'H1' };

  // Demo workstream group + sub-suite rows
  const demoData = [
    { group: 'PDM', suites: [
      { name: 'Iteration 2',                     depth: 0, pl: 120, ex: 95, pa: 80, fa: 15, ns: 25, ip: 0 },
      { name: '  Iteration 2.1',                 depth: 1, pl:  60, ex: 48, pa: 42, fa:  6, ns: 12, ip: 0 },
      { name: '    CPIMs',                       depth: 2, pl:  30, ex: 24, pa: 22, fa:  2, ns:  6, ip: 0 },
      { name: '      Phase 1',                   depth: 3, pl:  15, ex: 12, pa: 11, fa:  1, ns:  3, ip: 0 },
    ]},
  ];

  let rowIdx = 2;
  for (const { group, suites } of demoData) {
    const gRow = ws.getRow(rowIdx);
    const gCell = gRow.getCell(1);
    gCell.value = group;
    gCell.font  = font(true, 12, WHITE);
    gCell.fill  = fill(ACCENT);
    ws.mergeCells(`A${rowIdx}:H${rowIdx}`);
    rowIdx++;

    suites.forEach(({ name, pl, ex, pa, fa, ns, ip }, idx) => {
      const row    = ws.getRow(rowIdx);
      const passP  = ex ? Math.round((pa / ex) * 100) : 0;
      const rowFill = idx % 2 === 0 ? ALT_ROW : null;
      dataCell(row.getCell(1), name, { fillArgb: rowFill });
      dataCell(row.getCell(2), pl,   { fillArgb: rowFill, align: 'center' });
      dataCell(row.getCell(3), ex,   { fillArgb: rowFill, align: 'center' });
      dataCell(row.getCell(4), pa,   { fillArgb: pa > 0 ? GREEN_FILL : rowFill, align: 'center' });
      dataCell(row.getCell(5), fa,   { fillArgb: fa > 0 ? RED_FILL   : rowFill, align: 'center' });
      dataCell(row.getCell(6), ns,   { fillArgb: rowFill, align: 'center' });
      dataCell(row.getCell(7), ip,   { fillArgb: rowFill, align: 'center' });
      const pf = passP >= 80 ? GREEN_FILL : passP >= 60 ? YELLOW_FILL : RED_FILL;
      dataCell(row.getCell(8), `${passP}%`, { fillArgb: pf, align: 'center' });
      rowIdx++;
    });
  }

  ws.getRow(rowIdx + 1).getCell(1).value = '↑ The extension replaces this data on each run, grouped by workstream.';
  ws.getRow(rowIdx + 1).getCell(1).font  = { name: 'Calibri', size: 10, italic: true, color: { argb: MUTED } };
  ws.mergeCells(`A${rowIdx + 1}:H${rowIdx + 1}`);
}

// ─── Sheet 4: Bug Analysis ────────────────────────────────────────────────────

function buildBugAnalysis(wb) {
  const ws = wb.addWorksheet('Bug Analysis');
  widths(ws, [30, 14, 14]);
  freeze(ws);

  let r = 1;

  const sections = [
    {
      title: 'bugsBySeverity',
      rows: [['1 - Critical', 3, 8], ['2 - High', 12, 32], ['3 - Medium', 18, 49], ['4 - Low', 4, 11]],
      total: 37,
    },
    {
      title: 'bugsByPriority',
      rows: [['Priority 1', 5, 14], ['Priority 2', 15, 41], ['Priority 3', 12, 32], ['Priority 4', 5, 13]],
      total: 37,
    },
  ];

  for (const { title, rows, total } of sections) {
    const hCell = ws.getCell(r, 1);
    hCell.value = title;
    hCell.font  = font(true, 12, WHITE);
    hCell.fill  = fill(ACCENT);
    ws.mergeCells(`A${r}:C${r}`);
    r++;

    ['Value', 'Count', '%'].forEach((h, i) => headerCell(ws.getCell(r, i + 1), h));
    r++;

    rows.forEach(([label, count, pct], idx) => {
      dataCell(ws.getCell(r, 1), label,    { fillArgb: idx % 2 === 0 ? ALT_ROW : null });
      dataCell(ws.getCell(r, 2), count,    { fillArgb: idx % 2 === 0 ? ALT_ROW : null, align: 'center' });
      dataCell(ws.getCell(r, 3), `${pct}%`,{ fillArgb: idx % 2 === 0 ? ALT_ROW : null, align: 'center' });
      r++;
    });

    const totalRow = ws.getRow(r);
    dataCell(totalRow.getCell(1), 'TOTAL', { bold: true, fillArgb: HEADER_FILL });
    dataCell(totalRow.getCell(2), total,   { bold: true, fillArgb: HEADER_FILL, align: 'center' });
    dataCell(totalRow.getCell(3), '100%',  { bold: true, fillArgb: HEADER_FILL, align: 'center' });
    r += 2;
  }

  ws.getCell(`A${r}`).value = '↑ The extension replaces this data on each run based on your configured groupByFields.';
  ws.getCell(`A${r}`).font  = { name: 'Calibri', size: 10, italic: true, color: { argb: MUTED } };
  ws.mergeCells(`A${r}:C${r}`);
}

// ─── Sheet 5: Charts Guide ────────────────────────────────────────────────────

function buildChartsGuide(wb) {
  const ws = wb.addWorksheet('Charts Guide');
  widths(ws, [70]);

  const title = ws.getCell('A1');
  title.value = 'How to Add Charts to This Template';
  title.font  = font(true, 16, ACCENT.slice(2));
  ws.getRow(1).height = 26;

  const lines = [
    '',
    'After the extension populates data on each run, add charts once and they will auto-update:',
    '',
    'RECOMMENDED CHARTS',
    '',
    'Sheet "By Workstream":',
    '  1. Select A1:H5 (headers + 4 workstream rows)',
    '  2. Insert → Charts → Clustered Bar (shows Planned/Executed/Passed/Failed side by side)',
    '  3. Or select A1:A5 + H1:H5 for a Pass Rate bar chart per workstream',
    '',
    'Sheet "Sub-Suites":',
    '  1. Select the rows for one workstream block',
    '  2. Insert → Stacked Bar to show depth breakdown',
    '',
    'Sheet "Bug Analysis":',
    '  1. Select Value + Count columns in any section (e.g. A3:B6)',
    '  2. Insert → Pie or Donut chart for severity/priority distribution',
    '',
    'TOKENS (replaced each run on the Summary sheet)',
    '',
    '  {{Date}}            — report date (from Variables tab)',
    '  {{PP}}              — reporting period',
    '  {{FP}}              — forecast period',
    '  {{TotalPlanned}}    — consolidated planned test count',
    '  {{TotalExecuted}}   — consolidated executed test count',
    '  {{TotalPassed}}     — consolidated passed test count',
    '  {{TotalFailed}}     — consolidated failed test count',
    '  {{PassRate}}        — overall pass rate %',
    '  {{TotalNotStarted}} — consolidated not-started count',
    '  {{TotalInProgress}} — consolidated in-progress count',
    '  {{AISummary}}       — AI-generated executive summary (if ai-summary extension enabled)',
    '',
    'Add any custom {{TOKEN}} to the Summary sheet and map it in the Variables tab.',
    '',
    'TIP: Set EXCEL_TEMPLATE in your .env (or Credentials tab) to point to a customized copy.',
  ];

  lines.forEach((line, i) => {
    const cell = ws.getCell(i + 2, 1);
    cell.value = line;
    const isSection = line.startsWith('RECOMMENDED') || line.startsWith('TOKENS') || line.startsWith('Sheet') || line.startsWith('  {{');
    cell.font = line.startsWith('  {{')
      ? { name: 'Courier New', size: 10, color: { argb: 'FF374151' } }
      : font(line.startsWith('RECOMMENDED') || line.startsWith('TOKENS'), 11);
    cell.alignment = { wrapText: true };
  });
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main() {
  const wb = new ExcelJS.Workbook();
  wb.creator  = 'MMO Report Generator';
  wb.created  = new Date();
  wb.modified = new Date();

  buildSummary(wb);
  buildByWorkstream(wb);
  buildSubSuites(wb);
  buildBugAnalysis(wb);
  buildChartsGuide(wb);

  await wb.xlsx.writeFile(OUTPUT);
  console.log(`\nTemplate created: ${OUTPUT}`);
  console.log('\nNext steps:');
  console.log('  1. Open template.xlsx in Excel and add charts to "By Workstream", "Sub-Suites", "Bug Analysis" sheets');
  console.log('  2. Customize colors, layout, or add extra sheets as needed');
  console.log('  3. Map tokens in the Variables tab (auto-generate will create most of them)');
  console.log('  4. Enable "excel" in Output Formats — the extension loads this file each run\n');
}

main().catch(err => { console.error('Error:', err.message); process.exit(1); });
