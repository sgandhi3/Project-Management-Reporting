// One-time template fix: slide4's own "Grand Total" row (in the Priority
// Benefits per-plan breakdown table) was wired to BENEPBTTC-family tokens
// (Priority-only), even though the "SIT – 2026 Benefits" row above it adds
// more. Retokenize ONLY that Grand Total row to a new ACTIVEBEN* family that
// will sum Priority + any active non-Priority plans.
import fs from 'fs';
import PizZip from 'pizzip';

const TEMPLATE = process.argv[2] || 'temp2.pptx';
const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));
const SLIDE = 'ppt/slides/slide4.xml';
let xml = zip.file(SLIDE).asText();

const rows = [...xml.matchAll(/<a:tr h="\d+">[\s\S]*?<\/a:tr>/g)];
const grandTotalMatch = rows.find(m => m[0].includes('Grand Total') && m[0].includes('BENEPBTTC'));
if (!grandTotalMatch) {
  console.error('Grand Total row with BENEPBTTC not found — aborting, no changes made.');
  process.exit(1);
}

const OLD_TO_NEW = {
  '{{BENEPBTTC}}': '{{ACTIVEBENTTC}}',
  '{{BENEPBETC}}': '{{ACTIVEBENETC}}',
  '{{BENEPBPTC}}': '{{ACTIVEBENPTC}}',
  '{{BENEPBFTC}}': '{{ACTIVEBENFTC}}',
  '{{BENEPBIPTC}}': '{{ACTIVEBENIPTC}}',
  '{{BENEPBBTC}}': '{{ACTIVEBENBTC}}',
  '{{BENEPBNSTC}}': '{{ACTIVEBENNSTC}}',
  '{{BENEPBB}}': '{{ACTIVEBENB}}',
};

let newRow = grandTotalMatch[0];
for (const [oldTok, newTok] of Object.entries(OLD_TO_NEW)) {
  newRow = newRow.split(oldTok).join(newTok);
}

xml = xml.slice(0, grandTotalMatch.index) + newRow + xml.slice(grandTotalMatch.index + grandTotalMatch[0].length);
zip.file(SLIDE, xml);
fs.writeFileSync(TEMPLATE, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
console.log('Retokenized slide4 Grand Total row → ACTIVEBEN* tokens.');
