// One-time template fixes for the Benefits active-suite filtering
// inconsistency (MMO-specific, see the mmo-benefits-active-suite-fix branch):
//
// 1. Slide 4's own "Grand Total" row (in the Priority Benefits per-plan
//    breakdown table) was wired to BENEPBTTC-family tokens (Priority-only),
//    even though the "SIT – 2026 Benefits" row above it adds more.
//    Retokenized to a new ACTIVEBEN* family that sums Priority + any active
//    non-Priority plans (extensions/_dynamic-benefits.js adds the actual
//    per-plan rows to this table at runtime).
//
// 2. Slide 2's executive summary only had a "Priority Benefits - SIT" row —
//    once a non-Priority plan went active, its numbers had nowhere to show
//    up even though they were already counted in Grand Total. Rather than
//    adding a second row (and the layout reflow that would require), this
//    row is retokenized in place to the same ACTIVEBEN* combined total and
//    relabeled "2026 Benefits - SIT" (the plan program's name, not just its
//    first-phase "Priority" subset).
import fs from 'fs';
import PizZip from 'pizzip';

const TEMPLATE = process.argv[2] || 'temp2.pptx';
const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));

function retokenizeRow(xml, { slideLabel, matchText, requireToken, renameFrom, renameTo, tokenMap }) {
  const rows = [...xml.matchAll(/<a:tr h="\d+">[\s\S]*?<\/a:tr>/g)];
  const match = rows.find(m => m[0].includes(matchText) && (!requireToken || m[0].includes(requireToken)));
  if (!match) {
    console.log(`  (${slideLabel}) already applied or row not found — skipping`);
    return xml;
  }
  let newRow = match[0];
  if (renameFrom) newRow = newRow.split(`>${renameFrom}<`).join(`>${renameTo}<`);
  for (const [oldTok, newTok] of Object.entries(tokenMap)) {
    newRow = newRow.split(oldTok).join(newTok);
  }
  console.log(`  (${slideLabel}) retokenized`);
  return xml.slice(0, match.index) + newRow + xml.slice(match.index + match[0].length);
}

// ── Fix 1: slide4 Grand Total row ───────────────────────────────────────────
let slide4 = zip.file('ppt/slides/slide4.xml').asText();
slide4 = retokenizeRow(slide4, {
  slideLabel: 'slide4 Grand Total',
  matchText: 'Grand Total',
  requireToken: 'BENEPBTTC',
  tokenMap: {
    '{{BENEPBTTC}}': '{{ACTIVEBENTTC}}',
    '{{BENEPBETC}}': '{{ACTIVEBENETC}}',
    '{{BENEPBPTC}}': '{{ACTIVEBENPTC}}',
    '{{BENEPBFTC}}': '{{ACTIVEBENFTC}}',
    '{{BENEPBIPTC}}': '{{ACTIVEBENIPTC}}',
    '{{BENEPBBTC}}': '{{ACTIVEBENBTC}}',
    '{{BENEPBNSTC}}': '{{ACTIVEBENNSTC}}',
    '{{BENEPBB}}': '{{ACTIVEBENB}}',
  },
});
zip.file('ppt/slides/slide4.xml', slide4);

// ── Fix 2: slide2 "Priority Benefits - SIT" row → combined "2026 Benefits - SIT" ──
let slide2 = zip.file('ppt/slides/slide2.xml').asText();
slide2 = retokenizeRow(slide2, {
  slideLabel: 'slide2 Benefits row',
  matchText: 'Priority Benefits - SIT',
  renameFrom: 'Priority Benefits - SIT',
  renameTo: '2026 Benefits - SIT',
  tokenMap: {
    '{{BENEPBTTC}}':  '{{ACTIVEBENTTC}}',
    '{{BENEPBETC}}':  '{{ACTIVEBENETC}}',
    '{{BENEPBEP}}':   '{{ACTIVEBENEP}}',
    '{{BENEPBPTC}}':  '{{ACTIVEBENPTC}}',
    '{{BENEPBPP}}':   '{{ACTIVEBENPP}}',
    '{{BENEPBFTC}}':  '{{ACTIVEBENFTC}}',
    '{{BENEPBFP}}':   '{{ACTIVEBENFP}}',
    '{{BENEPBIPTC}}': '{{ACTIVEBENIPTC}}',
    '{{BENEPBB}}':    '{{ACTIVEBENB}}',
  },
});
zip.file('ppt/slides/slide2.xml', slide2);

fs.writeFileSync(TEMPLATE, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
console.log(`Done — ${TEMPLATE} updated.`);
