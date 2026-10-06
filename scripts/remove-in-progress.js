// One-time template patch: removes "In Progress" from every execution-stats
// table and from the Executive Summary KPI stat-box row, now that
// in-progress (and paused) test cases are folded into "Not Started"
// everywhere (see variables.js). Pure PizZip/regex surgery — see
// scripts/_pptx-xml-utils.js for why this avoids python-pptx.
//
// Usage: node scripts/remove-in-progress.js [path-to-template.pptx]
//   Defaults to MMO_Report_Template.pptx in the project root. Writes in place.

import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import { matchTopLevelTags, removeTableColumn } from './_pptx-xml-utils.js';

const TEMPLATE = process.argv[2] || path.join(process.cwd(), 'MMO_Report_Template.pptx');
const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));

function findGraphicFrameBlock(xml, shapeName) {
  const nameIdx = xml.indexOf(`name="${shapeName}"`);
  if (nameIdx === -1) return null;
  const start = xml.lastIndexOf('<p:graphicFrame>', nameIdx);
  const end = xml.indexOf('</p:graphicFrame>', nameIdx) + '</p:graphicFrame>'.length;
  if (start === -1 || end === -1) return null;
  return { start, end, text: xml.slice(start, end) };
}

function removeInProgressFromTable(xml, shapeName, label) {
  const frame = findGraphicFrameBlock(xml, shapeName);
  if (!frame) {
    console.log(`    ${label}: shape "${shapeName}" not found — skipping`);
    return xml;
  }
  const tblMatch = frame.text.match(/<a:tbl>[\s\S]*<\/a:tbl>/);
  if (!tblMatch) return xml;

  const newTbl = removeTableColumn(tblMatch[0], 'In Progress');
  if (!newTbl) {
    console.log(`    ${label}: no "In Progress" column found — skipping`);
    return xml;
  }
  const newFrameText = frame.text.slice(0, tblMatch.index) + newTbl + frame.text.slice(tblMatch.index + tblMatch[0].length);
  console.log(`    ${label}: removed "In Progress" column`);
  return xml.slice(0, frame.start) + newFrameText + xml.slice(frame.end);
}

const TABLE_TARGETS = [
  { slide: 'ppt/slides/slide2.xml',  label: 'Executive Summary',   tables: ['WorkstreamSummary'] },
  { slide: 'ppt/slides/slide3.xml',  label: 'PDM Execution',       tables: ['ExecSummary', 'WorkstreamSummary'] },
  { slide: 'ppt/slides/slide5.xml',  label: 'Benefits Execution',  tables: ['ExecSummary', 'WorkstreamSummary'] },
  { slide: 'ppt/slides/slide7.xml',  label: 'Enrollment Execution',tables: ['ExecSummary', 'WorkstreamSummary'] },
  { slide: 'ppt/slides/slide8.xml',  label: 'EDI Execution',       tables: ['ExecSummary', 'WorkstreamSummary'] },
];

for (const { slide, label, tables } of TABLE_TARGETS) {
  const file = zip.file(slide);
  if (!file) { console.warn(`  ⚠  ${slide} not found — skipping`); continue; }
  let xml = file.asText();
  console.log(`${label}:`);
  for (const t of tables) xml = removeInProgressFromTable(xml, t, `${label} / ${t}`);
  zip.file(slide, xml);
}

// ── Executive Summary KPI stat boxes ────────────────────────────────────────
// Delete the "In Progress" trio (StatBox autoshape + Num + Lbl) and evenly
// redistribute the remaining 5 across the original 6-box width.
const EXEC_SLIDE = 'ppt/slides/slide2.xml';
let execXml = zip.file(EXEC_SLIDE).asText();

const sps = matchTopLevelTags(execXml, 'p:sp');
function spInfo(sp) {
  const nameMatch = sp.text.match(/name="([^"]*)"/);
  const idMatch = sp.text.match(/<p:cNvPr id="(\d+)"/);
  const xfrmMatch = sp.text.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/);
  // NOTE: `sp` already has a `.text` field (the raw XML chunk, from
  // matchTopLevelTags) — the extracted plain-text label below MUST use a
  // different key, or the spread below silently overwrites the raw XML with
  // just the label string and any later `.text`-based edit corrupts the shape.
  const label = [...sp.text.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => m[1]).join('');
  return {
    ...sp,
    name: nameMatch ? nameMatch[1] : '',
    id: idMatch ? idMatch[1] : '',
    x: xfrmMatch ? Number(xfrmMatch[1]) : null,
    y: xfrmMatch ? Number(xfrmMatch[2]) : null,
    w: xfrmMatch ? Number(xfrmMatch[3]) : null,
    h: xfrmMatch ? Number(xfrmMatch[4]) : null,
    label,
  };
}
const allShapes = sps.map(spInfo);

const KPI_BAND_TOP = 1200000, KPI_BAND_BOTTOM = 2500000;
const inBand = s => s.y !== null && s.y > KPI_BAND_TOP && s.y < KPI_BAND_BOTTOM;

const lbl = allShapes.find(s => s.name.startsWith('Stat') && s.label.trim() === 'In Progress');
if (!lbl) {
  console.log('\n⚠  Could not find the "In Progress" KPI label — skipping KPI box fix');
} else {
  const trio = allShapes.filter(s => inBand(s) && Math.abs(s.x - lbl.x) < 50000);
  console.log(`\nDeleting "In Progress" KPI box trio (${trio.length} shapes)`);
  // Remove back-to-front so earlier offsets stay valid.
  const toRemove = [...trio].sort((a, b) => b.start - a.start);
  for (const s of toRemove) execXml = execXml.slice(0, s.start) + execXml.slice(s.end);

  // Re-scan after deletion for the remaining StatBox autoshapes.
  const spsAfter = matchTopLevelTags(execXml, 'p:sp').map(spInfo);
  const statBoxes = spsAfter.filter(s => s.name.startsWith('StatBox')).sort((a, b) => a.x - b.x);
  if (statBoxes.length !== 5) {
    console.log(`  ⚠  expected 5 remaining StatBox shapes, found ${statBoxes.length} — leaving positions as-is`);
  } else {
    const boxW = statBoxes[0].w;
    const spanStart = statBoxes[0].x;
    const spanEnd = statBoxes[4].x + statBoxes[4].w;
    const gap = (spanEnd - spanStart - 5 * boxW) / 4;

    for (let i = 0; i < 5; i++) {
      const box = statBoxes[i];
      const newX = Math.round(spanStart + i * (boxW + gap));
      const delta = newX - box.x;
      if (delta === 0) continue;
      const spsNow = matchTopLevelTags(execXml, 'p:sp').map(spInfo);
      const movingTrio = spsNow.filter(s => inBand(s) && Math.abs(s.x - box.x) < 50000)
        .sort((a, b) => b.start - a.start); // back-to-front
      for (const s of movingTrio) {
        const newText = s.text.replace(/<a:off x="(-?\d+)"/, (_, x) => `<a:off x="${Number(x) + delta}"`);
        execXml = execXml.slice(0, s.start) + newText + execXml.slice(s.end);
      }
      console.log(`  moved trio at x=${box.x} by ${delta} EMU -> x=${box.x + delta}`);
    }
  }
}

zip.file(EXEC_SLIDE, execXml);

fs.writeFileSync(TEMPLATE, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
console.log(`\n${TEMPLATE} updated.`);
