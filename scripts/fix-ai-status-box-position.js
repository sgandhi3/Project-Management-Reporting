// One-time template fix: the single-line AI status text box near the top of
// each workstream slide (PDM/Benefits/Enrollment/EDI) sat at a different
// vertical position on each slide — a pre-existing inconsistency from the
// original source template, not something introduced by the {{AI_...}}
// tokenization. Aligns Benefits/Enrollment/EDI's box to PDM's Y position
// (the one the user confirmed is correct). Only Y is changed — X/width/
// height are left as each slide's own, since only vertical position was
// reported as inconsistent.
import fs from 'fs';
import PizZip from 'pizzip';

const TEMPLATE = process.argv[2] || 'temp2.pptx';
const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));

const REFERENCE_SLIDE = 'ppt/slides/slide3.xml';
const REFERENCE_TOKEN = 'AI_PDM_CURSORY_STATUS';
const TARGETS = [
  { slide: 'ppt/slides/slide4.xml', token: 'AI_BENEFITS_STATUS' },
  { slide: 'ppt/slides/slide5.xml', token: 'AI_ENROLLMENT_STATUS' },
  { slide: 'ppt/slides/slide6.xml', token: 'AI_EDI_STATUS' },
];

function findShapeXfrm(xml, token) {
  const idx = xml.indexOf(token);
  if (idx === -1) return null;
  const spStart = xml.lastIndexOf('<p:sp>', idx);
  const spSnippet = xml.slice(spStart, spStart + 700);
  const m = spSnippet.match(/<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/);
  if (!m) return null;
  return { spStart, matchText: m[0], x: m[1], y: m[2], cx: m[3], cy: m[4], offsetInSnippet: spSnippet.indexOf(m[0]) };
}

const refXml = zip.file(REFERENCE_SLIDE).asText();
const ref = findShapeXfrm(refXml, REFERENCE_TOKEN);
if (!ref) {
  console.error(`Reference shape (${REFERENCE_TOKEN} on ${REFERENCE_SLIDE}) not found — aborting.`);
  process.exit(1);
}
console.log(`Reference Y (slide3, ${REFERENCE_TOKEN}): ${ref.y}`);

for (const { slide, token } of TARGETS) {
  let xml = zip.file(slide).asText();
  const shape = findShapeXfrm(xml, token);
  if (!shape) {
    console.warn(`  ⚠  ${slide}: ${token} shape not found — skipping`);
    continue;
  }
  if (shape.y === ref.y) {
    console.log(`  ${slide}: already at y=${ref.y} — skipping`);
    continue;
  }
  const newMatchText = shape.matchText.replace(/y="\d+"/, `y="${ref.y}"`);
  const absoluteIdx = shape.spStart + shape.offsetInSnippet;
  xml = xml.slice(0, absoluteIdx) + newMatchText + xml.slice(absoluteIdx + shape.matchText.length);
  zip.file(slide, xml);
  console.log(`  ${slide}: moved ${token} from y=${shape.y} to y=${ref.y}`);
}

fs.writeFileSync(TEMPLATE, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
console.log(`Done — ${TEMPLATE} updated.`);
