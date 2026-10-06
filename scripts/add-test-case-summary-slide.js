// One-time template patch: adds a new "Test Case Planning Summary" slide —
// one table, one row per workstream (WS | Test Cases in ADO | Test case
// count (Industry Best Practice) | Planned Execution | SIT End Date |
// Actual Execution | Comments) — positioned right after the Executive
// Summary slide. Mirrors the report owner's own hand-built example.
//
// All columns except Comments are driven by {{<WS>TCS...}} tokens computed
// in variables.js (config.js SIT_SCHEDULE / INDUSTRY_BENCHMARK_TEST_CASES
// supply the two inputs that aren't in ADO). Comments is left blank for
// manual fill-in in PowerPoint each week — see variables.js's TCSCOMMENTS
// tokens.
//
// This adds a brand-new slide PART (slide11.xml) — it does not renumber or
// touch slide1.xml..slide10.xml, so it's safe to run even though display
// order (presentation.xml's sldIdLst) and physical slide-file numbering
// already diverge in this deck (see extensions/_burndown-charts.js's
// CHART_TARGETS comment).
//
// Usage: node scripts/add-test-case-summary-slide.js [path-to-template.pptx]
//   Defaults to MMO_Report_Template.pptx in the project root. Writes in place.

import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import { titleRow, headerRow, dataRow, buildTable } from './_pptx-table-builder.js';

const TEMPLATE = process.argv[2] || path.join(process.cwd(), 'MMO_Report_Template.pptx');
const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));

const SLIDE_PART = 'ppt/slides/slide11.xml';
if (zip.file(SLIDE_PART)) {
  console.log(`${SLIDE_PART} already exists — skipping (already applied).`);
  process.exit(0);
}

// ── Build the table ─────────────────────────────────────────────────────────
const X = 501652, CX = 11163826;
const COL_WIDTHS = [1005840, 1234440, 1737360, 1280160, 1097280, 1280160, 3528586];
const HEADERS = ['WS', 'Test Cases in ADO', 'Test case count (Industry Best Practice)', 'Planned Execution', 'SIT End Date', 'Actual Execution', 'Comments'];

const WS_ROWS = [
  { label: 'PDM',        prefix: 'PDM' },
  { label: 'Benefits',   prefix: 'BENE' },
  { label: 'Enrollment', prefix: 'ENRO' },
  { label: 'EDI',        prefix: 'EDI' },
];

const rowsXml = [
  titleRow('Test Case Planning Summary — By Workstream', HEADERS.length, 230000),
  headerRow(HEADERS, 420000),
  ...WS_ROWS.map(({ label, prefix }) => dataRow([
    label,
    `{{${prefix}TCSADO}}`,
    `{{${prefix}TCSBENCH}}`,
    `{{${prefix}TCSPLANNED}}`,
    `{{${prefix}TCSSITEND}}`,
    `{{${prefix}TCSACTUAL}}`,
    `{{${prefix}TCSCOMMENTS}}`,
  ], { height: 280000 })),
];

const TABLE_CY = 230000 + 420000 + 280000 * WS_ROWS.length;
const TABLE_Y = 1490472; // 1.63in — matches the ExecSummary table y used on every other content slide

const table = buildTable({
  id: 4, name: 'TestCaseSummary',
  x: X, y: TABLE_Y, cx: CX, cy: TABLE_CY,
  colWidths: COL_WIDTHS, rowsXml,
});

const title = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="501652" y="280000"/><a:ext cx="11188700" cy="380000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/><a:lstStyle/><a:p><a:pPr marL="0" marR="0" lvl="0" indent="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1"><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="2400" b="1" kern="1200" dirty="0"><a:solidFill><a:prstClr val="black"/></a:solidFill><a:latin typeface="Calibri"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:rPr><a:t>Test Case Planning Summary | SIT Overview</a:t></a:r></a:p></p:txBody></p:sp>`;

const asOf = `<p:sp><p:nvSpPr><p:cNvPr id="3" name="AsOf"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="501652" y="680000"/><a:ext cx="11188700" cy="260000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US" sz="1200" b="1" u="sng" dirty="0"><a:solidFill><a:srgbClr val="313131"/></a:solidFill></a:rPr><a:t>As-of: {{Date}}</a:t></a:r></a:p></p:txBody></p:sp>`;

const footnote = `<p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="501652" y="${TABLE_Y + TABLE_CY + 120000}"/><a:ext cx="11188700" cy="260000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US" sz="900" i="1" dirty="0"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:rPr><a:t>Industry Best Practice counts are planning estimates (config.js), not a published external standard. Comments are filled in manually.</a:t></a:r></a:p></p:txBody></p:sp>`;

const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${title}${asOf}${table}${footnote}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

zip.file(SLIDE_PART, slideXml);

// ── Slide rels (slideLayout36, same as every other content slide; no notesSlide) ──
zip.file('ppt/slides/_rels/slide11.xml.rels',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout36.xml"/></Relationships>`);

// ── [Content_Types].xml — register the new slide part ──────────────────────
let contentTypes = zip.file('[Content_Types].xml').asText();
if (!contentTypes.includes('/ppt/slides/slide11.xml')) {
  contentTypes = contentTypes.replace(
    '</Types>',
    `<Override PartName="/ppt/slides/slide11.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`
  );
  zip.file('[Content_Types].xml', contentTypes);
}

// ── presentation.xml.rels — add the relationship PowerPoint resolves rIds through ──
let presRels = zip.file('ppt/_rels/presentation.xml.rels').asText();
const existingRelIds = [...presRels.matchAll(/Id="rId(\d+)"/g)].map(m => Number(m[1]));
const newRelId = `rId${Math.max(...existingRelIds) + 1}`;
presRels = presRels.replace(
  '</Relationships>',
  `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide11.xml"/></Relationships>`
);
zip.file('ppt/_rels/presentation.xml.rels', presRels);

// ── presentation.xml — insert into sldIdLst right after the Executive Summary slide ──
let presentation = zip.file('ppt/presentation.xml').asText();
const existingSldIds = [...presentation.matchAll(/<p:sldId id="(\d+)"/g)].map(m => Number(m[1]));
const newSldId = Math.max(...existingSldIds) + 1;

// rId7 = slide2.xml = Executive Summary (see ppt/_rels/presentation.xml.rels) —
// confirmed by direct inspection when this script was authored.
const EXEC_SUMMARY_SLDID_MARKER = /<p:sldId id="2147483620" r:id="rId7"\/>/;
if (!EXEC_SUMMARY_SLDID_MARKER.test(presentation)) {
  console.error('❌  Could not find the Executive Summary sldId entry (id=2147483620, r:id=rId7) in presentation.xml — the deck structure has changed since this script was written. Aborting without writing.');
  process.exit(1);
}
presentation = presentation.replace(
  EXEC_SUMMARY_SLDID_MARKER,
  `<p:sldId id="2147483620" r:id="rId7"/><p:sldId id="${newSldId}" r:id="${newRelId}"/>`
);
zip.file('ppt/presentation.xml', presentation);

fs.writeFileSync(TEMPLATE, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
console.log(`${TEMPLATE} updated — added slide11.xml ("Test Case Planning Summary") as ${newRelId}/sldId ${newSldId}, displayed right after the Executive Summary slide.`);
