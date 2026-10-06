// One-time template patch (follow-up to scripts/add-defect-aging-tables.js):
// moves the "Defect Aging" table from the bottom-right corner to directly
// beneath "Testing Defect Summary by Severity" (left column), on each of the
// 4 workstream Defect Report slides. Gives it its own teal title bar row so
// it reads the same as every other table in the deck (no separate caption
// shape needed). Shrinks/moves the "Key Insights on Defects" box down to
// make room, preserving its original bottom edge.
//
// Also removes the "Industry Best Practice counts are planning estimates..."
// footnote from the Test Case Planning Summary slide.
//
// Pure PizZip/regex surgery — see scripts/_pptx-xml-utils.js for why this
// avoids python-pptx on this deck.
//
// Usage: node scripts/reposition-aging-table.js [path-to-template.pptx]

import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import { matchTopLevelTags } from './_pptx-xml-utils.js';
import { titleRow, headerRow, dataRow, buildTable } from './_pptx-table-builder.js';

const TEMPLATE = process.argv[2] || path.join(process.cwd(), 'MMO_Report_Template.pptx');
const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));

function findBlockByName(xml, tag, shapeName) {
  const nameIdx = xml.indexOf(`name="${shapeName}"`);
  if (nameIdx === -1) return null;
  const start = xml.lastIndexOf(`<${tag}>`, nameIdx);
  const end = xml.indexOf(`</${tag}>`, nameIdx) + `</${tag}>`.length;
  if (start === -1 || end === -1) return null;
  return { start, end, text: xml.slice(start, end) };
}

function xfrmOf(blockText) {
  const m = blockText.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/);
  if (!m) return null;
  return { x: Number(m[1]), y: Number(m[2]), cx: Number(m[3]), cy: Number(m[4]) };
}

const DEFECT_SLIDES = [
  { slide: 'ppt/slides/slide4.xml',  prefix: 'PDM' },
  { slide: 'ppt/slides/slide6.xml',  prefix: 'BENE' },
  { slide: 'ppt/slides/slide9.xml',  prefix: 'ENRO' },
  { slide: 'ppt/slides/slide10.xml', prefix: 'EDI' },
];

const TITLE_H = 190000, HEADER_H = 200000, DATA_H = 250000;
const AGING_H = TITLE_H + HEADER_H + DATA_H;
const GAP_ABOVE = 60000, GAP_BELOW = 100000;

for (const { slide, prefix } of DEFECT_SLIDES) {
  const file = zip.file(slide);
  if (!file) { console.warn(`  ⚠  ${slide} not found — skipping`); continue; }
  let xml = file.asText();

  const sev = findBlockByName(xml, 'p:graphicFrame', 'SeverityTable');
  const oldCaption = findBlockByName(xml, 'p:sp', 'AgingCaption');
  const oldTable = findBlockByName(xml, 'p:graphicFrame', 'AgingTable');

  // Find the "Key Insights on Defects" NarrLeft box specifically (there are
  // two NarrLeft boxes per slide — this one, and the burndown-insights one).
  let narrDefects = null;
  for (const sp of matchTopLevelTags(xml, 'p:sp')) {
    if (sp.text.includes('name="NarrLeft"') && sp.text.includes('Key Insights on Defects')) {
      narrDefects = sp;
      break;
    }
  }

  if (!sev || !oldTable || !narrDefects) {
    console.warn(`  ⚠  ${slide}: missing SeverityTable/AgingTable/NarrLeft-Defects — skipping`);
    continue;
  }

  const sevXfrm = xfrmOf(sev.text);
  const narrXfrm = xfrmOf(narrDefects.text);
  const sevBottom = sevXfrm.y + sevXfrm.cy;
  const narrBottom = narrXfrm.y + narrXfrm.cy; // preserve this exact bottom edge

  const newTableY = sevBottom + GAP_ABOVE;
  const newNarrY = newTableY + AGING_H + GAP_BELOW;
  const newNarrH = narrBottom - newNarrY;

  if (newNarrH < 600000) {
    console.warn(`  ⚠  ${slide}: computed insights-box height too small (${newNarrH} EMU) — skipping`);
    continue;
  }

  // Remove old caption (if present) + old table — back-to-front by position.
  const toRemove = [oldCaption, oldTable].filter(Boolean).sort((a, b) => b.start - a.start);
  for (const block of toRemove) xml = xml.slice(0, block.start) + xml.slice(block.end);

  // Resize/move the NarrLeft box (re-locate it fresh since offsets may have shifted).
  let narrNow = null;
  for (const sp of matchTopLevelTags(xml, 'p:sp')) {
    if (sp.text.includes('name="NarrLeft"') && sp.text.includes('Key Insights on Defects')) { narrNow = sp; break; }
  }
  const newNarrText = narrNow.text.replace(
    /<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/,
    (_, x) => `<a:off x="${x}" y="${newNarrY}"/><a:ext cx="${narrXfrm.cx}" cy="${newNarrH}"/>`
  );
  xml = xml.slice(0, narrNow.start) + newNarrText + xml.slice(narrNow.end);

  // Build + insert the new 3-row Aging table (title bar + header + data).
  const colW = Math.floor(sevXfrm.cx / 4);
  const colWidths = [colW, colW, colW, sevXfrm.cx - colW * 3];
  const rows = [
    titleRow('Open Defects by Age (Days Since Logged)', 4, TITLE_H),
    headerRow(['0-10 Days', '11-20 Days', '21-30 Days', '31+ Days'], HEADER_H),
    dataRow([`{{${prefix}AGE010}}`, `{{${prefix}AGE1120}}`, `{{${prefix}AGE2130}}`, `{{${prefix}AGE31P}}`], { height: DATA_H, boldFirst: false }),
  ];
  const newTable = buildTable({
    id: 9300, name: 'AgingTable',
    x: sevXfrm.x, y: newTableY, cx: sevXfrm.cx, cy: AGING_H,
    colWidths, rowsXml: rows,
  });
  xml = xml.replace('</p:spTree>', `${newTable}</p:spTree>`);

  zip.file(slide, xml);
  console.log(`  ${slide} (${prefix}): aging table moved under SeverityTable (y=${newTableY}), Key Insights on Defects now y=${newNarrY} h=${newNarrH}`);
}

// ── Remove the Test Case Planning Summary slide's footnote ─────────────────
const TCS_SLIDE = 'ppt/slides/slide11.xml';
const tcsFile = zip.file(TCS_SLIDE);
if (tcsFile) {
  let tcsXml = tcsFile.asText();
  let removed = false;
  for (const sp of matchTopLevelTags(tcsXml, 'p:sp')) {
    if (sp.text.includes('Industry Best Practice counts are planning estimates')) {
      tcsXml = tcsXml.slice(0, sp.start) + tcsXml.slice(sp.end);
      removed = true;
      break;
    }
  }
  if (removed) {
    zip.file(TCS_SLIDE, tcsXml);
    console.log('  Test Case Planning Summary: removed footnote text box');
  } else {
    console.log('  ⚠  Test Case Planning Summary: footnote text box not found — skipping');
  }
}

fs.writeFileSync(TEMPLATE, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
console.log(`\n${TEMPLATE} updated.`);
