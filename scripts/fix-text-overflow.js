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
// AI-narrative sentences. Matched by name/content rather than hardcoded shape
// IDs, since IDs shift whenever the template's shapes are edited/reordered —
// the stat-box names ("Stat1Num" etc.) and the narrative sentence's own
// wording are what stay stable across template edits.
const TARGET_PREDICATES = {
  2: (name, _text) => /^NarrLeft$/.test(name) || /Num$/.test(name), // NarrLeft, Stat1Num..Stat4Num (x2)
  3: (_name, text) => /planned test cases/i.test(text),             // PDM status sentence
  4: (_name, text) => /planned test cases/i.test(text),             // EDI status sentence
  5: (_name, text) => /planned test cases/i.test(text),             // Enrollment status sentence
};

function fixShapeSegment(shapeXml) {
  // Self-closing bodyPr with no autofit child: <a:bodyPr .../>
  const selfClosing = shapeXml.match(/<a:bodyPr([^>]*)\/>/);
  // bodyPr with children (e.g. <a:spAutoFit/> or <a:noAutofit/>) already present
  const withChildren = shapeXml.match(/<a:bodyPr([^>]*)>([\s\S]*?)<\/a:bodyPr>/);

  if (withChildren) {
    const [full, attrs, inner] = withChildren;
    if (inner.includes('<a:normAutofit')) {
      return { seg: shapeXml, changed: false, reason: 'already normAutofit' };
    }
    const newInner = inner.replace(/<a:(spAutoFit|noAutofit)\/>/, '<a:normAutofit/>');
    const finalInner = newInner.includes('<a:normAutofit')
      ? newInner
      : `<a:normAutofit/>${inner}`; // no recognized autofit child — add one
    return { seg: shapeXml.replace(full, `<a:bodyPr${attrs}>${finalInner}</a:bodyPr>`), changed: true };
  } else if (selfClosing) {
    const [full, attrs] = selfClosing;
    return { seg: shapeXml.replace(full, `<a:bodyPr${attrs}><a:normAutofit/></a:bodyPr>`), changed: true };
  }
  return { seg: shapeXml, changed: false, reason: 'no bodyPr found' };
}

// Scans every <p:sp> shape in document order, applying fixShapeSegment to
// the ones the predicate selects, and rebuilds the slide XML from the
// original text so unmatched shapes pass through untouched.
function fixSlide(xml, predicate) {
  const matches = [];
  let cursor = 0;
  let out = '';
  for (const m of xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const seg = m[0];
    const cnv = seg.match(/<p:cNvPr id="(\d+)" name="([^"]*)"/);
    let segToUse = seg;
    if (cnv) {
      const [, id, name] = cnv;
      const text = [...seg.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(x => x[1]).join('');
      if (predicate(name, text)) {
        const fixed = fixShapeSegment(seg);
        matches.push({ id, name, changed: fixed.changed, reason: fixed.reason });
        if (fixed.changed) segToUse = fixed.seg;
      }
    }
    out += xml.slice(cursor, m.index) + segToUse;
    cursor = m.index + seg.length;
  }
  out += xml.slice(cursor);
  return { xml: out, matches };
}

function main() {
  if (!fs.existsSync(TEMPLATE)) {
    console.error(`Template not found: ${TEMPLATE}`);
    process.exit(1);
  }
  const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));
  let totalChanged = 0, totalSkipped = 0;

  for (const [slide, predicate] of Object.entries(TARGET_PREDICATES)) {
    const name = `ppt/slides/slide${slide}.xml`;
    const file = zip.file(name);
    if (!file) { console.warn(`⚠  ${name} not found in ${TEMPLATE}`); continue; }

    const { xml, matches } = fixSlide(file.asText(), predicate);
    console.log(`slide${slide}:`);
    if (matches.length === 0) {
      console.log('  ⚠ no matching shapes found — template may have changed');
    }
    for (const m of matches) {
      if (m.changed) {
        console.log(`  ✓ shape id=${m.id} name=${m.name} → normAutofit set`);
        totalChanged++;
      } else {
        console.log(`  · shape id=${m.id} name=${m.name} — ${m.reason}`);
        totalSkipped++;
      }
    }
    zip.file(name, xml);
  }

  fs.writeFileSync(TEMPLATE, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log(`\n${totalChanged} shape(s) updated, ${totalSkipped} already-fine/not-found, in ${TEMPLATE}`);
}

main();
