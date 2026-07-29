// Excel output extension
//
// Loads template.xlsx, replaces {{TOKEN}} placeholders in every cell with
// live data values, repopulates the data sheets (By Workstream, Sub-Suites,
// Bug Analysis), then writes the populated workbook to the output path.
//
// This mirrors how ppt.js works: the template owns the layout/styling/charts,
// the extension only fills in the data.
//
// Config:
//   EXCEL_TEMPLATE — path to the template file (default: ./template.xlsx)
//   --out <path>   — output path (default: Report_YYYY-MM-DD.xlsx)
//
// To create the starter template: node create-excel-template.js

import ExcelJS from 'exceljs';
import fs from 'fs';
import { existsSync, readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { VARIABLE_MAP } from '../variables.js';

const __dirname2     = path.dirname(fileURLToPath(import.meta.url));
const UI_CONFIG_PATH = path.join(__dirname2, '..', 'ui-config.json');

const args      = process.argv.slice(2);
const getArg    = flag => { const i = args.indexOf(flag); return i !== -1 ? args[i + 1] : null; };
const TEMPLATE  = process.env.EXCEL_TEMPLATE || path.join(process.cwd(), 'template.xlsx');
const outputArg = getArg('--out') || path.join(process.cwd(), `Report_${new Date().toISOString().slice(0, 10)}.xlsx`);

// ─── Variable resolution (same as ppt.js) ─────────────────────────────────────

function buildReplacements(data) {
  let effectiveMap = VARIABLE_MAP;

  if (existsSync(UI_CONFIG_PATH)) {
    const uiConfig  = JSON.parse(readFileSync(UI_CONFIG_PATH, 'utf8'));
    const uiMappings = uiConfig.variableMappings || [];
    if (uiMappings.length > 0) {
      effectiveMap = {};
      for (const { token, path: expr } of uiMappings) {
        if (!token) continue;
        effectiveMap[token] = new Function('d', `try { return ${expr}; } catch { return ''; }`);
      }
    }
  }

  const replacements = {};
  for (const [key, getter] of Object.entries(effectiveMap)) {
    try {
      replacements[`{{${key}}}`] = String(getter(data) ?? '');
    } catch {
      replacements[`{{${key}}}`] = '';
    }
  }
  return replacements;
}

// ─── Token replacement ────────────────────────────────────────────────────────

function replaceTokensInSheet(sheet, replacements) {
  sheet.eachRow({ includeEmpty: false }, row => {
    row.eachCell({ includeEmpty: false }, cell => {
      if (typeof cell.value === 'string' && cell.value.includes('{{')) {
        let val = cell.value;
        for (const [token, replacement] of Object.entries(replacements)) {
          val = val.split(token).join(replacement);
        }
        cell.value = val;
      }
    });
  });
}

// ─── Style helpers for data rows ──────────────────────────────────────────────

const GREEN  = 'FFD1FAE5';
const RED    = 'FFFEE2E2';
const YELLOW = 'FFFEF3C7';
const ALT    = 'FFF8FAFC';
const ACCENT = 'FF2563EB';
const WHITE  = 'FFFFFFFF';
const HFILL  = 'FFE0E7FF';

function applyFill(cell, argb) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function applyBorder(cell) {
  const t = { style: 'thin', color: { argb: 'FFD1D5DB' } };
  cell.border = { top: t, left: t, bottom: t, right: t };
}

function setDataCell(cell, value, { bold = false, fillArgb = null, align = 'left' } = {}) {
  cell.value = value;
  cell.font  = { name: 'Calibri', size: 11, bold };
  cell.alignment = { vertical: 'middle', horizontal: align };
  applyBorder(cell);
  if (fillArgb) applyFill(cell, fillArgb);
}

function setHeaderCell(cell, text) {
  cell.value = text;
  cell.font  = { name: 'Calibri', size: 12, bold: true, color: { argb: WHITE } };
  cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: ACCENT } };
  cell.border = { top: { style: 'thin', color: { argb: 'FFD1D5DB' } }, left: { style: 'thin', color: { argb: 'FFD1D5DB' } }, bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } }, right: { style: 'thin', color: { argb: 'FFD1D5DB' } } };
  cell.alignment = { vertical: 'middle', horizontal: 'center' };
}

// Unmerge every merged range that touches rows fromRow..toRow, then re-merge safely
function clearMergesInRange(sheet, fromRow, toRow) {
  const toRemove = [];
  sheet._merges && Object.keys(sheet._merges).forEach(key => {
    const m = sheet._merges[key];
    // m.model has { top, bottom, left, right }
    const model = m.model || m;
    if (model.top >= fromRow && model.bottom <= toRow) toRemove.push(key);
  });
  for (const key of toRemove) {
    try { sheet.unMergeCells(key); } catch { /* ignore */ }
  }
}

function safeMerge(sheet, range) {
  try { sheet.unMergeCells(range); } catch { /* already unmerged */ }
  sheet.mergeCells(range);
}

// ─── Clear data rows (rows 2+) in a sheet, preserving row 1 header ───────────

function clearDataRows(sheet, fromRow = 2, toRow = 200) {
  for (let r = fromRow; r <= toRow; r++) {
    const row = sheet.getRow(r);
    // Stop clearing once we hit genuinely empty rows well past our data
    let hasContent = false;
    row.eachCell({ includeEmpty: false }, () => { hasContent = true; });
    if (!hasContent && r > fromRow + 5) break;
    row.eachCell({ includeEmpty: true }, cell => {
      cell.value  = null;
      cell.fill   = { type: 'pattern', pattern: 'none' };
      cell.border = {};
      cell.font   = {};
    });
  }
}

// ─── Populate By Workstream sheet ─────────────────────────────────────────────

function populateWorkstreamSheet(wb, data) {
  const ws = wb.getWorksheet('By Workstream');
  if (!ws || !data.stats) return;

  clearDataRows(ws);

  const entries = Object.entries(data.stats);
  let totalPlan = 0, totalExec = 0, totalPass = 0, totalFail = 0, totalNS = 0, totalIP = 0, totalBlocked = 0;

  entries.forEach(([wsName, s], idx) => {
    const row     = ws.getRow(idx + 2);
    const passP   = s.executed ? Math.round((s.passed / s.executed) * 100) : 0;
    const rowFill = idx % 2 === 0 ? ALT : null;
    const blocked = s.blocked ?? 0;

    setDataCell(row.getCell(1), wsName,       { fillArgb: rowFill });
    setDataCell(row.getCell(2), s.planned,    { fillArgb: rowFill,                        align: 'center' });
    setDataCell(row.getCell(3), s.executed,   { fillArgb: rowFill,                        align: 'center' });
    setDataCell(row.getCell(4), s.passed,     { fillArgb: s.passed  > 0 ? GREEN : rowFill, align: 'center' });
    setDataCell(row.getCell(5), s.failed,     { fillArgb: s.failed  > 0 ? RED   : rowFill, align: 'center' });
    setDataCell(row.getCell(6), s.notStarted, { fillArgb: rowFill,                        align: 'center' });
    setDataCell(row.getCell(7), s.inProgress, { fillArgb: rowFill,                        align: 'center' });
    setDataCell(row.getCell(8), blocked,      { fillArgb: blocked   > 0 ? YELLOW : rowFill, align: 'center' });
    const pf = passP >= 80 ? GREEN : passP >= 60 ? YELLOW : RED;
    setDataCell(row.getCell(9), `${passP}%`,  { fillArgb: pf, align: 'center' });

    totalPlan += s.planned; totalExec += s.executed;
    totalPass += s.passed;  totalFail += s.failed;
    totalNS   += s.notStarted; totalIP += s.inProgress; totalBlocked += blocked;
  });

  // Total row
  const totalRowIdx = entries.length + 2;
  const totalPassP  = totalExec ? Math.round((totalPass / totalExec) * 100) : 0;
  const totalRow    = ws.getRow(totalRowIdx);
  [['TOTAL', null], [totalPlan, null], [totalExec, null], [totalPass, GREEN], [totalFail, RED],
   [totalNS, null], [totalIP, null], [totalBlocked, totalBlocked > 0 ? YELLOW : null],
   [`${totalPassP}%`, totalPassP >= 80 ? GREEN : totalPassP >= 60 ? YELLOW : RED]]
    .forEach(([v, f], i) => {
      const cell = totalRow.getCell(i + 1);
      setDataCell(cell, v, { bold: true, fillArgb: f || HFILL, align: i === 0 ? 'left' : 'center' });
    });
}

// ─── Populate Sub-Suites sheet ────────────────────────────────────────────────

function populateSubSuitesSheet(wb, data) {
  const ws = wb.getWorksheet('Sub-Suites');
  if (!ws || !data.subStats) return;

  const hasAny = Object.values(data.subStats).some(v => Object.keys(v).length > 0);
  if (!hasAny) return;

  clearMergesInRange(ws, 2, 500);
  clearDataRows(ws);

  let rowIdx = 2;
  let altCount = 0;

  for (const [wsName, suites] of Object.entries(data.subStats)) {
    if (Object.keys(suites).length === 0) continue;

    // Workstream group header
    const hRow  = ws.getRow(rowIdx);
    const hCell = hRow.getCell(1);
    hCell.value = wsName;
    hCell.font  = { name: 'Calibri', size: 12, bold: true, color: { argb: WHITE } };
    hCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: ACCENT } };
    safeMerge(ws, `A${rowIdx}:H${rowIdx}`);
    rowIdx++;

    for (const [suitePath, s] of Object.entries(suites)) {
      const depth  = (suitePath.match(/ \/ /g) || []).length;
      const indent = '  '.repeat(depth);
      const label  = indent + suitePath.split(' / ').pop();
      const passP  = s.executed ? Math.round((s.passed / s.executed) * 100) : 0;
      const rowFill = altCount % 2 === 0 ? ALT : null;
      const row    = ws.getRow(rowIdx);

      const blocked = s.blocked ?? 0;
      setDataCell(row.getCell(1), label,        { bold: depth === 0, fillArgb: rowFill });
      setDataCell(row.getCell(2), s.planned,    { fillArgb: rowFill, align: 'center' });
      setDataCell(row.getCell(3), s.executed,   { fillArgb: rowFill, align: 'center' });
      setDataCell(row.getCell(4), s.passed,     { fillArgb: s.passed  > 0 ? GREEN  : rowFill, align: 'center' });
      setDataCell(row.getCell(5), s.failed,     { fillArgb: s.failed  > 0 ? RED    : rowFill, align: 'center' });
      setDataCell(row.getCell(6), s.notStarted, { fillArgb: rowFill, align: 'center' });
      setDataCell(row.getCell(7), s.inProgress, { fillArgb: rowFill, align: 'center' });
      setDataCell(row.getCell(8), blocked,      { fillArgb: blocked   > 0 ? YELLOW : rowFill, align: 'center' });
      const pf = passP >= 80 ? GREEN : passP >= 60 ? YELLOW : RED;
      setDataCell(row.getCell(9), `${passP}%`,  { fillArgb: pf, align: 'center' });

      altCount++;
      rowIdx++;
    }
  }
}

// ─── Populate Bug Analysis sheet ──────────────────────────────────────────────

function populateBugSheet(wb, data) {
  const ws = wb.getWorksheet('Bug Analysis');
  if (!ws) return;

  const groupByKeys = Object.keys(data).filter(k => /By[A-Z]/.test(k));
  if (groupByKeys.length === 0) return;

  clearMergesInRange(ws, 1, 300);
  clearDataRows(ws, 1, 300);

  let rowIdx = 1;

  for (const key of groupByKeys) {
    const val = data[key];
    if (!val || typeof val !== 'object') continue;

    const hCell = ws.getCell(rowIdx, 1);
    hCell.value = key;
    hCell.font  = { name: 'Calibri', size: 12, bold: true, color: { argb: WHITE } };
    hCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: ACCENT } };
    safeMerge(ws, `A${rowIdx}:C${rowIdx}`);
    rowIdx++;

    setHeaderCell(ws.getCell(rowIdx, 1), 'Value');
    setHeaderCell(ws.getCell(rowIdx, 2), 'Count');
    setHeaderCell(ws.getCell(rowIdx, 3), '%');
    rowIdx++;

    const { total, ...buckets } = val;
    const sorted = Object.entries(buckets).sort((a, b) => b[1] - a[1]);

    sorted.forEach(([label, count], idx) => {
      const pct     = total > 0 ? Math.round((count / total) * 100) : 0;
      const rowFill = idx % 2 === 0 ? ALT : null;
      setDataCell(ws.getCell(rowIdx, 1), label,    { fillArgb: rowFill });
      setDataCell(ws.getCell(rowIdx, 2), count,    { fillArgb: rowFill, align: 'center' });
      setDataCell(ws.getCell(rowIdx, 3), `${pct}%`,{ fillArgb: rowFill, align: 'center' });
      rowIdx++;
    });

    rowIdx++; // blank between sections
  }
}

// ─── Main generate function ───────────────────────────────────────────────────

export async function generate(data) {
  if (!fs.existsSync(TEMPLATE)) {
    console.error(`\n❌  Excel template not found: ${TEMPLATE}`);
    console.error('    Run: node create-excel-template.js');
    return;
  }

  const replacements = buildReplacements(data);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TEMPLATE);

  // Replace {{TOKEN}} in all cells across all sheets (Summary sheet primarily)
  wb.eachSheet(sheet => replaceTokensInSheet(sheet, replacements));

  // Repopulate dynamic data sheets with live data
  populateWorkstreamSheet(wb, data);
  populateSubSuitesSheet(wb, data);
  populateBugSheet(wb, data);

  await wb.xlsx.writeFile(outputArg);
  console.log(`\nExcel report saved → ${outputArg}`);

  data._generatedFiles = data._generatedFiles || [];
  data._generatedFiles.push(outputArg);
}
