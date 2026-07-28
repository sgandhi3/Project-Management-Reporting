// Excel output extension
//
// Creates a multi-sheet .xlsx workbook from the collected data.
//
// Config (via .env or CLI):
//   --out <path>  — where to write the output file (default: Report_YYYY-MM-DD.xlsx in cwd)
//
// Sheet 1: Summary       — key metrics, AI summary if present, variable token values
// Sheet 2: By Workstream — per-workstream stats with conditional formatting
// Sheet 3: Sub-Suites    — sub-suite breakdown (only if data exists)
// Sheet 4: Bug Analysis  — groupBy tables (only if groupBy keys exist)
// Sheet 5: Charts Guide  — instructions for adding charts in Excel

import ExcelJS from 'exceljs';
import fs from 'fs';
import { existsSync, readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { VARIABLE_MAP } from '../variables.js';

const __dirname2 = path.dirname(fileURLToPath(import.meta.url));
const UI_CONFIG_PATH = path.join(__dirname2, '..', 'ui-config.json');

const args      = process.argv.slice(2);
const getArg    = flag => { const i = args.indexOf(flag); return i !== -1 ? args[i + 1] : null; };
const outputArg = getArg('--out') || path.join(process.cwd(), `Report_${new Date().toISOString().slice(0, 10)}.xlsx`);

// ─── Color constants ──────────────────────────────────────────────────────────
const ACCENT       = '2563EB';
const WHITE        = 'FFFFFFFF';
const GREEN_FILL   = 'FFD1FAE5';
const RED_FILL     = 'FFFEE2E2';
const YELLOW_FILL  = 'FFFEF3C7';
const ALT_ROW      = 'FFF8FAFC';
const HEADER_FILL  = 'FFE0E7FF';

// ─── Style helpers ────────────────────────────────────────────────────────────

function applyFont(cell, bold = false, size = 11, color = null) {
  cell.font = { name: 'Calibri', size, bold, ...(color ? { color: { argb: color } } : {}) };
}

function applyFill(cell, argb) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function applyBorder(cell) {
  const thin = { style: 'thin', color: { argb: 'FFD1D5DB' } };
  cell.border = { top: thin, left: thin, bottom: thin, right: thin };
}

function applyHeaderStyle(cell, text) {
  cell.value = text;
  applyFont(cell, true, 12, WHITE);
  applyFill(cell, 'FF' + ACCENT);
  cell.alignment = { vertical: 'middle', horizontal: 'center' };
  applyBorder(cell);
}

function applyAltRow(row, colCount, even) {
  if (!even) return;
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    if (!cell.fill || cell.fill.fgColor?.argb === 'FFFFFFFF') {
      applyFill(cell, ALT_ROW);
    }
  }
}

function freezeAndFilter(sheet, ySplit = 1) {
  sheet.views = [{ state: 'frozen', ySplit }];
}

function setColWidths(sheet, widths) {
  widths.forEach((w, i) => {
    sheet.getColumn(i + 1).width = w;
  });
}

// ─── Variable resolution (same as ppt.js) ─────────────────────────────────────

function buildEffectiveMap(data) {
  let effectiveMap = VARIABLE_MAP;
  if (existsSync(UI_CONFIG_PATH)) {
    const uiConfig = JSON.parse(readFileSync(UI_CONFIG_PATH, 'utf8'));
    const uiMappings = uiConfig.variableMappings || [];
    if (uiMappings.length > 0) {
      effectiveMap = {};
      for (const { token, path: expr } of uiMappings) {
        if (!token) continue;
        effectiveMap[token] = new Function('d', `try { return ${expr}; } catch { return ''; }`);
      }
    }
  }
  const resolved = {};
  for (const [key, getter] of Object.entries(effectiveMap)) {
    try { resolved[key] = String(getter(data) ?? ''); }
    catch { resolved[key] = ''; }
  }
  return resolved;
}

// ─── Sheet 1: Summary ─────────────────────────────────────────────────────────

function buildSummarySheet(wb, data, resolved) {
  const ws = wb.addWorksheet('Summary');
  setColWidths(ws, [24, 18, 18, 18, 18, 18]);
  freezeAndFilter(ws);

  const today = new Date().toISOString().slice(0, 10);
  let row = 1;

  // Title row — merged across 6 cols
  const titleRow = ws.getRow(row);
  titleRow.height = 28;
  const titleCell = ws.getCell(`A${row}`);
  titleCell.value = `Weekly SIT Status Report — ${today}`;
  applyFont(titleCell, true, 16, WHITE);
  applyFill(titleCell, 'FF' + ACCENT);
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.mergeCells(`A${row}:F${row}`);
  row++;

  // Blank row
  row++;

  // KEY METRICS label
  const kmCell = ws.getCell(`A${row}`);
  kmCell.value = 'KEY METRICS';
  applyFont(kmCell, true, 12);
  row++;

  // Metric table header
  ['Metric', 'Value'].forEach((h, i) => {
    const cell = ws.getCell(row, i + 1);
    applyHeaderStyle(cell, h);
  });
  row++;

  const c = data.consolidatedData || {};
  const passRate = c.executed ? Math.round((c.passed / c.executed) * 100) : 0;
  const failRate = c.executed ? Math.round((c.failed / c.executed) * 100) : 0;
  const metrics = [
    ['Total Planned',  c.planned    ?? 0, null],
    ['Total Executed', c.executed   ?? 0, null],
    ['Total Passed',   c.passed     ?? 0, GREEN_FILL],
    ['Total Failed',   c.failed     ?? 0, RED_FILL],
    ['Pass Rate %',    `${passRate}%`,  null],
    ['Fail Rate %',    `${failRate}%`,  null],
  ];

  metrics.forEach(([label, value, fill], idx) => {
    const dataRow = ws.getRow(row);
    const labelCell = dataRow.getCell(1);
    const valueCell = dataRow.getCell(2);
    labelCell.value = label;
    valueCell.value = value;
    applyFont(labelCell, false, 11);
    applyFont(valueCell, true, 11);
    applyBorder(labelCell);
    applyBorder(valueCell);
    if (fill) {
      applyFill(labelCell, fill);
      applyFill(valueCell, fill);
    } else if (idx % 2 === 0) {
      applyFill(labelCell, ALT_ROW);
      applyFill(valueCell, ALT_ROW);
    }
    row++;
  });

  // AI Summary section
  if (data._aiSummary) {
    row++;
    const aiLabelCell = ws.getCell(`A${row}`);
    aiLabelCell.value = 'AI SUMMARY';
    applyFont(aiLabelCell, true, 12, ACCENT);
    row++;

    const aiCell = ws.getCell(`A${row}`);
    aiCell.value = data._aiSummary;
    aiCell.alignment = { wrapText: true, vertical: 'top' };
    applyFont(aiCell, false, 11);
    applyFill(aiCell, 'FFE0E7FF');
    ws.mergeCells(`A${row}:F${row}`);
    ws.getRow(row).height = Math.min(200, 15 * Math.ceil(data._aiSummary.length / 80));
    row++;
  }

  // Variable token values at bottom
  const tokenEntries = Object.entries(resolved).filter(([token]) => token);
  if (tokenEntries.length > 0) {
    row++;
    const vlCell = ws.getCell(`A${row}`);
    vlCell.value = 'VARIABLE TOKEN VALUES';
    applyFont(vlCell, true, 12);
    row++;

    ['Token', 'Value'].forEach((h, i) => {
      const cell = ws.getCell(row, i + 1);
      applyHeaderStyle(cell, h);
    });
    row++;

    tokenEntries.forEach(([token, value], idx) => {
      const dataRow = ws.getRow(row);
      const tCell = dataRow.getCell(1);
      const vCell = dataRow.getCell(2);
      tCell.value = `{{${token}}}`;
      vCell.value = value;
      applyFont(tCell, false, 10);
      applyFont(vCell, false, 10);
      applyBorder(tCell);
      applyBorder(vCell);
      if (idx % 2 === 0) {
        applyFill(tCell, ALT_ROW);
        applyFill(vCell, ALT_ROW);
      }
      row++;
    });
  }
}

// ─── Sheet 2: By Workstream ───────────────────────────────────────────────────

function buildWorkstreamSheet(wb, data) {
  const ws = wb.addWorksheet('By Workstream');
  const cols = ['Workstream', 'Planned', 'Executed', 'Passed', 'Failed', 'Not Started', 'In Progress', 'Pass %'];
  setColWidths(ws, [22, 12, 12, 12, 12, 14, 14, 12]);
  freezeAndFilter(ws);

  // Header
  cols.forEach((h, i) => applyHeaderStyle(ws.getCell(1, i + 1), h));
  ws.getRow(1).height = 22;

  const stats   = data.stats || {};
  const entries = Object.entries(stats);
  let totalPlan = 0, totalExec = 0, totalPass = 0, totalFail = 0, totalNS = 0, totalIP = 0;

  entries.forEach(([ws_name, s], idx) => {
    const row = ws.getRow(idx + 2);
    const passP = s.executed ? Math.round((s.passed / s.executed) * 100) : 0;
    const values = [ws_name, s.planned, s.executed, s.passed, s.failed, s.notStarted, s.inProgress, `${passP}%`];
    values.forEach((v, ci) => {
      const cell = row.getCell(ci + 1);
      cell.value = v;
      applyFont(cell, false, 11);
      applyBorder(cell);
      // Conditional formatting
      if (ci === 3 && s.passed > 0) applyFill(cell, GREEN_FILL);       // Passed
      else if (ci === 4 && s.failed > 0) applyFill(cell, RED_FILL);    // Failed
      else if (ci === 7) {                                               // Pass %
        if (passP >= 80) applyFill(cell, GREEN_FILL);
        else if (passP >= 60) applyFill(cell, YELLOW_FILL);
        else applyFill(cell, RED_FILL);
      } else if (idx % 2 === 0) {
        applyFill(cell, ALT_ROW);
      }
    });
    totalPlan += s.planned; totalExec += s.executed;
    totalPass += s.passed;  totalFail += s.failed;
    totalNS   += s.notStarted; totalIP += s.inProgress;
  });

  // Total row
  const totalRowIdx = entries.length + 2;
  const totalRow = ws.getRow(totalRowIdx);
  const totalPassPct = totalExec ? Math.round((totalPass / totalExec) * 100) : 0;
  const totals = ['TOTAL', totalPlan, totalExec, totalPass, totalFail, totalNS, totalIP, `${totalPassPct}%`];
  totals.forEach((v, i) => {
    const cell = totalRow.getCell(i + 1);
    cell.value = v;
    applyFont(cell, true, 11);
    applyFill(cell, HEADER_FILL);
    applyBorder(cell);
  });

  ws.autoFilter = { from: 'A1', to: `H1` };
  return entries.length;
}

// ─── Sheet 3: Sub-Suites ──────────────────────────────────────────────────────

function buildSubSuitesSheet(wb, data) {
  const subStats = data.subStats || {};
  const hasAny = Object.values(subStats).some(v => Object.keys(v).length > 0);
  if (!hasAny) return false;

  const ws = wb.addWorksheet('Sub-Suites');
  const cols = ['Suite', 'Planned', 'Executed', 'Passed', 'Failed', 'Not Started', 'In Progress', 'Pass %'];
  setColWidths(ws, [38, 12, 12, 12, 12, 14, 14, 12]);
  freezeAndFilter(ws);

  // Header
  cols.forEach((h, i) => applyHeaderStyle(ws.getCell(1, i + 1), h));
  ws.getRow(1).height = 22;

  let rowIdx = 2;
  let altCount = 0;

  for (const [wsName, suites] of Object.entries(subStats)) {
    if (Object.keys(suites).length === 0) continue;

    // Workstream header row
    const headerRow = ws.getRow(rowIdx);
    const headerCell = headerRow.getCell(1);
    headerCell.value = wsName;
    applyFont(headerCell, true, 12, WHITE);
    applyFill(headerCell, 'FF' + ACCENT);
    ws.mergeCells(`A${rowIdx}:H${rowIdx}`);
    rowIdx++;

    for (const [suitePath, s] of Object.entries(suites)) {
      const depth  = (suitePath.match(/ \/ /g) || []).length;
      const indent = '  '.repeat(depth);
      const label  = suitePath.split(' / ').pop();
      const passP  = s.executed ? Math.round((s.passed / s.executed) * 100) : 0;
      const values = [`${indent}${label}`, s.planned, s.executed, s.passed, s.failed, s.notStarted, s.inProgress, `${passP}%`];
      const row    = ws.getRow(rowIdx);

      values.forEach((v, ci) => {
        const cell = row.getCell(ci + 1);
        cell.value = v;
        applyFont(cell, ci === 0 && depth === 0, 11);
        applyBorder(cell);
        // Conditional formatting
        if (ci === 3 && s.passed > 0) applyFill(cell, GREEN_FILL);
        else if (ci === 4 && s.failed > 0) applyFill(cell, RED_FILL);
        else if (ci === 7) {
          if (passP >= 80) applyFill(cell, GREEN_FILL);
          else if (passP >= 60) applyFill(cell, YELLOW_FILL);
          else applyFill(cell, RED_FILL);
        } else if (altCount % 2 === 0) {
          applyFill(cell, ALT_ROW);
        }
      });
      altCount++;
      rowIdx++;
    }
  }

  ws.autoFilter = { from: 'A1', to: 'H1' };
  return true;
}

// ─── Sheet 4: Bug Analysis ────────────────────────────────────────────────────

function buildBugSheet(wb, data) {
  const groupByKeys = Object.keys(data).filter(k => /By[A-Z]/.test(k));
  if (groupByKeys.length === 0) return false;

  const ws = wb.addWorksheet('Bug Analysis');
  setColWidths(ws, [30, 14, 14]);
  freezeAndFilter(ws);

  let rowIdx = 1;

  for (const key of groupByKeys) {
    const val = data[key];
    if (!val || typeof val !== 'object') continue;

    // Section header
    const headerRow = ws.getRow(rowIdx);
    const headerCell = headerRow.getCell(1);
    headerCell.value = key;
    applyFont(headerCell, true, 12, WHITE);
    applyFill(headerCell, 'FF' + ACCENT);
    ws.mergeCells(`A${rowIdx}:C${rowIdx}`);
    rowIdx++;

    // Column headers
    ['Value', 'Count', '%'].forEach((h, i) => {
      applyHeaderStyle(ws.getCell(rowIdx, i + 1), h);
    });
    rowIdx++;

    const { total, ...buckets } = val;
    const sorted = Object.entries(buckets).sort((a, b) => b[1] - a[1]);

    sorted.forEach(([label, count], idx) => {
      const pct  = total > 0 ? Math.round((count / total) * 100) : 0;
      const row  = ws.getRow(rowIdx);
      [label, count, `${pct}%`].forEach((v, ci) => {
        const cell = row.getCell(ci + 1);
        cell.value = v;
        applyFont(cell, false, 11);
        applyBorder(cell);
        if (idx % 2 === 0) applyFill(cell, ALT_ROW);
      });
      rowIdx++;
    });

    rowIdx++; // blank row between sections
  }

  return true;
}

// ─── Sheet 5: Charts Guide ────────────────────────────────────────────────────

function buildChartsGuideSheet(wb, wsCount) {
  const ws = wb.addWorksheet('Charts Guide');
  setColWidths(ws, [60]);

  let row = 1;

  const titleCell = ws.getCell(`A${row}`);
  titleCell.value = 'How to Add Charts in Excel';
  applyFont(titleCell, true, 16, ACCENT);
  row += 2;

  const instructions = [
    'The data in sheets 2–4 is structured for charting. Follow these steps:',
    '',
    '1. Select the data range on the target sheet.',
    '2. Go to Insert → Charts → and pick the recommended chart type below.',
    '3. Customize titles, colors, and labels as needed.',
    '',
    'RECOMMENDED CHARTS PER SHEET',
    '',
    `Sheet 2 — "By Workstream":`,
    `  • Select A1:H${wsCount + 1} and insert a Clustered Bar chart.`,
    '  • This shows Planned vs Executed vs Passed vs Failed per workstream.',
    '  • Alternatively, select just the Workstream + Pass % columns for a simple pass-rate chart.',
    '',
    `Sheet 3 — "Sub-Suites":`,
    '  • Select the rows for a specific workstream block and insert a Stacked Bar chart.',
    '  • This visualizes test execution breakdown per sub-suite.',
    '',
    `Sheet 4 — "Bug Analysis":`,
    '  • Each section (bugsBy*) can be charted with a Pie or Donut chart.',
    '  • Select the Value and Count columns for any section and insert Pie chart.',
    '',
    'TIP: Use "Recommended Charts" (Insert → Recommended Charts) for automatic suggestions.',
  ];

  instructions.forEach(line => {
    const cell = ws.getCell(`A${row}`);
    cell.value = line;
    if (line.startsWith('RECOMMENDED') || line.startsWith('Sheet')) {
      applyFont(cell, true, 11);
    } else {
      applyFont(cell, false, 11);
    }
    cell.alignment = { wrapText: true };
    row++;
  });
}

// ─── Main generate function ───────────────────────────────────────────────────

export async function generate(data) {
  const resolved  = buildEffectiveMap(data);
  const wb        = new ExcelJS.Workbook();
  wb.creator      = 'MMO Report Generator';
  wb.created      = new Date();

  buildSummarySheet(wb, data, resolved);
  const wsCount = buildWorkstreamSheet(wb, data);
  buildSubSuitesSheet(wb, data);
  buildBugSheet(wb, data);
  buildChartsGuideSheet(wb, wsCount);

  await wb.xlsx.writeFile(outputArg);
  console.log(`\nExcel report saved → ${outputArg}`);

  data._generatedFiles = data._generatedFiles || [];
  data._generatedFiles.push(outputArg);
}
