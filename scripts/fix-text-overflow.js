// Prevents visual text overlap on shapes whose content length varies at
// runtime (token-substituted numbers, AI-written narrative sentences).
//
// Two failure modes were found in temp2.pptx, both letting overflowing text
// spill into neighboring shapes instead of staying inside their own box:
//   1. Stat1Num–Stat4Num and the exec-summary narrative box: <a:bodyPr> had
//      no autofit element at all — text just overflows the box boundary.
//   2. The per-workstream status sentences (slides 3–5): <a:bodyPr> had
//      <a:spAutoFit/> ("resize shape to fit text") — this GROWS the box
//      instead of shrinking the text, and since PowerPoint shapes are
//      absolutely positioned, a growing box overlaps whatever sits below it
//      rather than pushing it out of the way.
//
// The fix for both is the same: <a:normAutofit/> ("shrink text on
// overflow"). This keeps the shape's size and position exactly as designed
// — no style change — and only reduces font size when the actual content
// would otherwise overflow.
//
// Usage: node scripts/fix-text-overflow.js [path-to-pptx]  (default: temp2.pptx)
import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';

const TEMPLATE = process.argv[2] || path.join(process.cwd(), 'temp2.pptx');

// Shapes whose text length varies at runtime: token-substituted numbers or
// AI-narrative sentences. Update this list if new variable-content shapes
// are added to the template.
const TARGET_SHAPES = {
  2: [61, 29, 35, 38, 44, 47, 50], // NarrLeft, Stat1Num..Stat4Num (x2)
  3: [7],                          // PDM status sentence
  4: [6],                          // EDI status sentence
  5: [6],                          // Enrollment status sentence
};

function fixShape(xml, shapeId) {
  const idIdx = xml.indexOf(`<p:cNvPr id="${shapeId}"`);
  if (idIdx === -1) return { xml, changed: false, reason: 'shape not found' };
  const spStart = xml.lastIndexOf('<p:sp>', idIdx);
  const spEnd = xml.indexOf('</p:sp>', idIdx) + '</p:sp>'.length;
  if (spStart === -1 || spEnd === -1) return { xml, changed: false, reason: 'shape boundaries not found' };

  const shapeXml = xml.slice(spStart, spEnd);

  // Self-closing bodyPr with no autofit child: <a:bodyPr .../>
  const selfClosing = shapeXml.match(/<a:bodyPr([^>]*)\/>/);
  // bodyPr with children (e.g. <a:spAutoFit/> or <a:noAutofit/>) already present
  const withChildren = shapeXml.match(/<a:bodyPr([^>]*)>([\s\S]*?)<\/a:bodyPr>/);

  let newShapeXml;
  if (withChildren) {
    const [full, attrs, inner] = withChildren;
    if (inner.includes('<a:normAutofit')) {
      return { xml, changed: false, reason: 'already normAutofit' };
    }
    const newInner = inner.replace(/<a:(spAutoFit|noAutofit)\/>/, '<a:normAutofit/>');
    const finalInner = newInner.includes('<a:normAutofit')
      ? newInner
      : `<a:normAutofit/>${inner}`; // no recognized autofit child — add one
    newShapeXml = shapeXml.replace(full, `<a:bodyPr${attrs}>${finalInner}</a:bodyPr>`);
  } else if (selfClosing) {
    const [full, attrs] = selfClosing;
    newShapeXml = shapeXml.replace(full, `<a:bodyPr${attrs}><a:normAutofit/></a:bodyPr>`);
  } else {
    return { xml, changed: false, reason: 'no bodyPr found' };
  }

  return { xml: xml.slice(0, spStart) + newShapeXml + xml.slice(spEnd), changed: true };
}

function main() {
  if (!fs.existsSync(TEMPLATE)) {
    console.error(`Template not found: ${TEMPLATE}`);
    process.exit(1);
  }
  const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));
  let totalChanged = 0, totalSkipped = 0;

  for (const [slide, shapeIds] of Object.entries(TARGET_SHAPES)) {
    const name = `ppt/slides/slide${slide}.xml`;
    const file = zip.file(name);
    if (!file) { console.warn(`⚠  ${name} not found in ${TEMPLATE}`); continue; }

    let xml = file.asText();
    console.log(`slide${slide}:`);
    for (const shapeId of shapeIds) {
      const result = fixShape(xml, shapeId);
      xml = result.xml;
      if (result.changed) {
        console.log(`  ✓ shape id=${shapeId} → normAutofit set`);
        totalChanged++;
      } else {
        console.log(`  · shape id=${shapeId} — ${result.reason}`);
        totalSkipped++;
      }
    }
    zip.file(name, xml);
  }

  fs.writeFileSync(TEMPLATE, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log(`\n${totalChanged} shape(s) updated, ${totalSkipped} already-fine/not-found, in ${TEMPLATE}`);
}

main();
