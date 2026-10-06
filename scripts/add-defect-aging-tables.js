// One-time template patch: adds a compact "Defect Aging" table (open
// defects bucketed 0-10 / 11-20 / 21-30 / 31+ days since logged) to each
// workstream's Testing Defect Report slide, directly beneath the existing
// "Key Insights on <WS> Defect Burndown" text box. Counts are computed live
// by variables.js (extensions/_defect-aging.js) via {{<WS>AGE...}} tokens —
// this script only lays out the table shape once; ppt.js's normal
// {{TOKEN}} replacement pass fills it in on every run after this.
//
// Usage: node scripts/add-defect-aging-tables.js [path-to-template.pptx]
//   Defaults to MMO_Report_Template.pptx in the project root. Writes in place.

import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import { titleRow, headerRow, dataRow, buildTable, findShapeXfrm } from './_pptx-table-builder.js';

const TEMPLATE = process.argv[2] || path.join(process.cwd(), 'MMO_Report_Template.pptx');

const TARGETS = [
  { slide: 'ppt/slides/slide4.xml',  prefix: 'PDM',  burndownToken: 'AI_PDM_BURNDOWN_INSIGHTS' },
  { slide: 'ppt/slides/slide6.xml',  prefix: 'BENE', burndownToken: 'AI_BENEFITS_BURNDOWN_INSIGHTS' },
  { slide: 'ppt/slides/slide9.xml',  prefix: 'ENRO', burndownToken: 'AI_ENROLLMENT_BURNDOWN_INSIGHTS' },
  { slide: 'ppt/slides/slide10.xml', prefix: 'EDI',  burndownToken: 'AI_EDI_BURNDOWN_INSIGHTS' },
];

const GAP = 60000; // ~0.066in breathing room below the insights text box
const CAPTION_H = 190000; // ~0.21in
const TABLE_H = 460000; // ~0.50in (header + 1 data row)
const BOTTOM_MARGIN = 90000; // ~0.10in, kept clear of the slide edge

function captionSp(id, x, y, cx, cy, text) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="AgingCaption"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US" sz="1100" b="1" dirty="0"><a:solidFill><a:srgbClr val="313131"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));
let changed = 0;

for (const { slide, prefix, burndownToken } of TARGETS) {
  const file = zip.file(slide);
  if (!file) { console.warn(`  ⚠  ${slide} not found — skipping`); continue; }
  let xml = file.asText();

  if (xml.includes('AgingCaption') && xml.includes(`${prefix}AGE010`)) {
    console.log(`  ${slide}: aging table already present — skipping`);
    continue;
  }

  const anchor = findShapeXfrm(xml, burndownToken);
  if (!anchor) {
    console.warn(`  ⚠  ${slide}: couldn't locate ${burndownToken} box — skipping`);
    continue;
  }

  const capY = anchor.y + anchor.cy + GAP;
  const tableY = capY + CAPTION_H;
  // Keep the block on-slide even if the anchor box runs unexpectedly low.
  const SLIDE_H = 6858000;
  const maxTableH = SLIDE_H - BOTTOM_MARGIN - tableY;
  const tableH = Math.min(TABLE_H, Math.max(maxTableH, 300000));

  const colW = Math.floor(anchor.cx / 4);
  const colWidths = [colW, colW, colW, anchor.cx - colW * 3];

  const rows = [
    headerRow(['0-10 Days', '11-20 Days', '21-30 Days', '31+ Days'], 120000),
    dataRow(
      [`{{${prefix}AGE010}}`, `{{${prefix}AGE1120}}`, `{{${prefix}AGE2130}}`, `{{${prefix}AGE31P}}`],
      { height: tableH - 120000, boldFirst: false }
    ),
  ];

  const caption = captionSp(9101, anchor.x, capY, anchor.cx, CAPTION_H, 'Open Defects by Age (Days Since Logged):');
  const table = buildTable({
    id: 9102, name: 'AgingTable',
    x: anchor.x, y: tableY, cx: anchor.cx, cy: tableH,
    colWidths, rowsXml: rows,
  });

  xml = xml.replace('</p:spTree>', `${caption}${table}</p:spTree>`);
  zip.file(slide, xml);
  changed++;
  console.log(`  ${slide}: added Defect Aging caption + table for ${prefix} at y=${capY}`);
}

if (changed > 0) {
  fs.writeFileSync(TEMPLATE, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log(`\n${TEMPLATE} updated (${changed} slide(s)).`);
} else {
  console.log('\nNo changes made.');
}
