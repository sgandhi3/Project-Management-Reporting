// One-time template patch: rebuilds the Test Case Planning Summary table
// (slide11.xml) with the revised column set — drops "Test case count
// (Industry Best Practice)" and adds Planned Execution %, Actual Execution
// %, Delta #, and a colored Status cell (see extensions/_tcs-status.js,
// which paints that cell at generation time — the template only lays down a
// neutral placeholder fill + the {{<WS>TCSSTATUS}} locator token).
//
// Simplest to fully rebuild rather than surgically insert/remove columns in
// the existing table — same graphicFrame position, new structure. Pure
// PizZip/regex surgery (see scripts/_pptx-xml-utils.js for why).
//
// Usage: node scripts/rebuild-test-case-summary-table.js [path-to-template.pptx]

import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import { titleRow, headerRow, customRow, buildTable } from './_pptx-table-builder.js';

const TEMPLATE = process.argv[2] || path.join(process.cwd(), 'MMO_Report_Template.pptx');
const SLIDE = 'ppt/slides/slide11.xml';

const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));
const file = zip.file(SLIDE);
if (!file) { console.error(`${SLIDE} not found — aborting.`); process.exit(1); }
let xml = file.asText();

const nameIdx = xml.indexOf('name="TestCaseSummary"');
if (nameIdx === -1) { console.error('TestCaseSummary table not found — aborting.'); process.exit(1); }
const start = xml.lastIndexOf('<p:graphicFrame>', nameIdx);
const end = xml.indexOf('</p:graphicFrame>', nameIdx) + '</p:graphicFrame>'.length;
const oldFrame = xml.slice(start, end);
const oldXfrm = oldFrame.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/);
const [, x, y, cx] = oldXfrm;

const HEADERS = ['WS', 'Test Cases in ADO', 'Planned Execution', 'Planned Execution %', 'SIT End Date', 'Actual Execution', 'Actual Execution %', 'Delta #', 'Status', 'Comments'];
const COL_WIDTHS = [822960, 1005840, 914400, 914400, 914400, 914400, 914400, 822960, 548640, 3391426];

const WS_ROWS = [
  { label: 'PDM',        prefix: 'PDM' },
  { label: 'Benefits',   prefix: 'BENE' },
  { label: 'Enrollment', prefix: 'ENRO' },
  { label: 'EDI',        prefix: 'EDI' },
];

const TITLE_H = 230000, HEADER_H = 420000, DATA_H = 280000;

const rowsXml = [
  titleRow('Test Case Planning Summary — By Workstream', HEADERS.length, TITLE_H),
  headerRow(HEADERS, HEADER_H),
  ...WS_ROWS.map(({ label, prefix }) => customRow([
    { text: label, bold: true },
    { text: `{{${prefix}TCSADO}}` },
    { text: `{{${prefix}TCSPLANNED}}` },
    { text: `{{${prefix}TCSPLANNEDPCT}}%` },
    { text: `{{${prefix}TCSSITEND}}` },
    { text: `{{${prefix}TCSACTUAL}}` },
    { text: `{{${prefix}TCSACTUALPCT}}%` },
    { text: `{{${prefix}TCSDELTA}}` },
    { text: `{{${prefix}TCSSTATUS}}`, fill: 'BDBDBD' },
    { text: `{{${prefix}TCSCOMMENTS}}` },
  ], DATA_H)),
];

const newCy = TITLE_H + HEADER_H + DATA_H * WS_ROWS.length;
const newFrame = buildTable({
  id: 4, name: 'TestCaseSummary',
  x, y, cx, cy: newCy,
  colWidths: COL_WIDTHS, rowsXml,
});

xml = xml.slice(0, start) + newFrame + xml.slice(end);
zip.file(SLIDE, xml);

fs.writeFileSync(TEMPLATE, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
console.log(`${TEMPLATE} updated — TestCaseSummary table rebuilt with ${HEADERS.length} columns.`);
