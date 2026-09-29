// One-time migration utility: converts the report owner's hand-built
// New_Temp_source.pptx (a real filled-in example deck, no {{TOKEN}}
// placeholders) into the tokenized template the weekly pipeline actually
// runs against (extensions/ppt.js reads PPTX_TEMPLATE from .env).
//
// This is NOT part of the weekly run — it's a build step you re-run only if
// the report owner hands over a revised static mockup and wants it
// retokenized. It works by literal text-run POSITION: for each slide it
// walks every <a:t>...</a:t> run in document order (0-indexed) and replaces
// the ones listed below with the matching {{TOKEN}}, leaving every other
// run (headers, legends, footnotes) untouched. Position-based rather than
// string-based specifically so repeated literal values (lots of cells just
// say "0") don't collide.
//
// If the source deck's layout changes at all — a row added/removed/
// reordered, a sentence split into an extra run — these indices will no
// longer line up and this script must be re-derived by re-dumping the
// source's <a:t> runs in order (see git history/PR description for the
// method used to derive the maps below).
//
// Usage: node scripts/build-template-from-source.js
//   Reads:  New_Temp_source.pptx (project root)
//   Writes: MMO_Report_Template.pptx (project root) — point PPTX_TEMPLATE
//           at this file.

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import PizZip from 'pizzip';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SOURCE = path.join(ROOT, 'New_Temp_source.pptx');
const OUTPUT = path.join(ROOT, 'MMO_Report_Template.pptx');

// index -> replacement string. Only indices listed here are touched;
// everything else in the slide is left exactly as-is. An empty string
// blanks a run (used when a multi-run sentence is consolidated into a
// single token on its first run).
const SLIDE_MAPS = {
  'ppt/slides/slide1.xml': {
    1: '{{Date}}',
  },

  'ppt/slides/slide2.xml': {
    3: '{{AI_OVERALL_STATUS}}',
    4: '{{AI_PDM_ITERATION_UPDATE}}',
    5: '{{AI_PDM_DEFECT_SUMMARY}}',
    6: '{{AI_BENEFITS_UPDATE}}',
    7: '',
    9: '{{AI_DEFECT_TRIAGE}}',
    15: 'As-of: {{Date}}',
    // PDM - Cursory
    30: '{{PDMCTTC}}', 31: '{{PDMCETC}}', 32: '{{PDMCEP}}%', 33: '{{PDMCPTC}}',
    34: '{{PDMCPP}}%', 35: '{{PDMCFTC}}', 36: '{{PDMCFP}}%', 37: '{{PDMCIPTC}}', 38: '{{PDMCURSORYDEFOPEN}}',
    // PDM - SIT
    40: '{{PDMSITTC}}', 41: '{{PDMSITETC}}', 42: '{{PDMSITEP}}%', 43: '{{PDMSITPTC}}',
    44: '{{PDMSITPP}}%', 45: '{{PDMSITFTC}}', 46: '{{PDMSITFP}}%', 47: '{{PDMSITIPTC}}', 48: '{{PDMSITDEFOPEN}}',
    // 2026 Benefits - SIT (= all active benefits)
    50: '{{ACTIVEBENTTC}}', 51: '{{ACTIVEBENETC}}', 52: '{{ACTIVEBENEP}}%', 53: '{{ACTIVEBENPTC}}',
    54: '{{ACTIVEBENPP}}%', 55: '{{ACTIVEBENFTC}}', 56: '{{ACTIVEBENFP}}%', 57: '{{ACTIVEBENIPTC}}', 58: '{{ACTIVEBENB}}',
    // Enrollment - SIT
    60: '{{ENROTTC}}', 61: '{{ENROETC}}', 62: '{{ENROEP}}%', 63: '{{ENROPTC}}',
    64: '{{ENROPP}}%', 65: '{{ENROFTC}}', 66: '{{ENROFP}}%', 67: '{{ENROIPTC}}', 68: '{{ENROB}}',
    // EDI - SIT
    70: '{{EDITTC}}', 71: '{{EDIETC}}', 72: '{{EDIEP}}%', 73: '{{EDIPTC}}',
    74: '{{EDIPP}}%', 75: '{{EDIFTC}}', 76: '{{EDIFP}}%', 77: '{{EDIIPTC}}', 78: '{{EDIB}}',
    // Grand Total
    80: '{{TTC}}', 81: '{{ETC}}', 82: '{{EP}}%', 83: '{{PTC}}',
    84: '{{PP}}%', 85: '{{FTC}}', 86: '{{FP}}%', 87: '{{IPTC}}', 88: '{{TB}}',
    // KPI cards
    97: '{{TTC}}*', 99: '{{PTC}}', 100: 'Passed ({{PP}}%)', 101: '{{FTC}}',
    102: 'Failed ({{FP}}%)', 103: '{{TB}}', 105: '{{ETC}}', 107: '{{IPTC}}',
  },

  'ppt/slides/slide3.xml': {
    2: 'As-of: {{Date}}',
    // Cursory / Iteration 2
    20: '{{PDMIT2TTC}}', 21: '{{PDMIT2ETC}}', 22: '{{PDMIT2EP}}%', 23: '{{PDMIT2PTC}}',
    24: '{{PDMIT2PP}}%', 25: '{{PDMIT2FTC}}', 26: '{{PDMIT2FP}}%', 27: '{{PDMIT2IPTC}}',
    28: '{{PDMIT2BTC}}', 29: '{{PDMIT2NSTC}}', 30: '{{PDMIT2DEFOPEN}}',
    // Cursory / Iteration 2.1
    33: '{{PDMIT21TTC}}', 34: '{{PDMIT21ETC}}', 35: '{{PDMIT21EP}}%', 36: '{{PDMIT21PTC}}',
    37: '{{PDMIT21PP}}%', 38: '{{PDMIT21FTC}}', 39: '{{PDMIT21FP}}%', 40: '{{PDMIT21IPTC}}',
    41: '{{PDMIT21BTC}}', 42: '{{PDMIT21NSTC}}', 43: '{{PDMIT21DEFOPEN}}',
    // Cursory / Iteration 3
    46: '{{PDMIT3TTC}}', 47: '{{PDMIT3ETC}}', 48: '{{PDMIT3EP}}%', 49: '{{PDMIT3PTC}}',
    50: '{{PDMIT3PP}}%', 51: '{{PDMIT3FTC}}', 52: '{{PDMIT3FP}}%', 53: '{{PDMIT3IPTC}}',
    54: '{{PDMIT3BTC}}', 55: '{{PDMIT3NSTC}}', 56: '{{PDMIT3DEFOPEN}}',
    // SIT
    59: '{{PDMSITTC}}', 60: '{{PDMSITETC}}', 61: '{{PDMSITEP}}%', 62: '{{PDMSITPTC}}',
    63: '{{PDMSITPP}}%', 64: '{{PDMSITFTC}}', 65: '{{PDMSITFP}}%', 66: '{{PDMSITIPTC}}',
    67: '{{PDMSITBTC}}', 68: '{{PDMSITNSTC}}', 69: '{{PDMSITDEFOPEN}}',
    // Status paragraph
    88: '{{AI_PDM_CURSORY_STATUS}}', 89: '', 90: '',
    // SIT Review Summary sub-table — PDM row
    104: '{{PDMSITTC}}', 106: '{{PDMSITETC}}', 107: '{{PDMSITPTC}}', 108: '{{PDMSITFTC}}',
    109: '{{PDMSITIPTC}}', 110: '{{PDMSITBTC}}', 111: '{{PDMSITNSTC}}', 112: '{{PDMSITDEFOPEN}}',
    // Grand Total
    114: '{{PDMSITTC}}', 116: '{{PDMSITETC}}', 117: '{{PDMSITPTC}}', 118: '{{PDMSITFTC}}',
    119: '{{PDMSITIPTC}}', 120: '{{PDMSITBTC}}', 121: '{{PDMSITNSTC}}', 122: '{{PDMCB}}*',
  },

  'ppt/slides/slide4.xml': {
    2: 'As-of: {{Date}}',
    10: '{{PDMIT2DEFTOTAL}}', 11: '{{PDMIT2DEFOPEN}}', 12: '{{PDMIT2DEFCLOSED}}', 13: '{{PDMIT2DEFRESOLVED}}',
    15: '{{PDMIT21DEFTOTAL}}', 16: '{{PDMIT21DEFOPEN}}', 17: '{{PDMIT21DEFCLOSED}}', 18: '{{PDMIT21DEFRESOLVED}}',
    20: '{{PDMIT3DEFTOTAL}}', 21: '{{PDMIT3DEFOPEN}}', 22: '{{PDMIT3DEFCLOSED}}', 23: '{{PDMIT3DEFRESOLVED}}',
    25: '{{PDMSITDEFTOTAL}}', 26: '{{PDMSITDEFOPEN}}', 27: '{{PDMSITDEFCLOSED}}', 28: '{{PDMSITDEFRESOLVED}}',
    30: '{{PDMDEFTOTAL}}', 31: '{{PDMCB}}', 32: '{{PDMDEFCLOSEDTOTAL}}', 33: '{{PDMDEFRESOLVEDTOTAL}}',
    40: '{{AI_PDM_DEFECT_STATUS}}', 41: '', 42: '', 43: '', 44: '',
    47: '{{AI_PDM_DEFECT_INSIGHTS}}', 48: '', 49: '', 50: '', 51: '', 52: '', 53: '', 54: '', 55: '', 56: '', 57: '', 58: '',
    61: '{{AI_PDM_BURNDOWN_INSIGHTS}}', 62: '', 63: '', 64: '', 65: '',
  },

  'ppt/slides/slide5.xml': {
    6: 'As-of: {{Date}}',
    // SIT - Priority Benefit
    23: '{{BENEPBTTC}}', 24: '{{BENEPBETC}}', 25: '{{BENEPBEP}}%', 26: '{{BENEPBPTC}}',
    27: '{{BENEPBPP}}%', 28: '{{BENEPBFTC}}', 29: '{{BENEPBFP}}%', 30: '{{BENEPBIPTC}}',
    31: '{{BENEPBBTC}}', 32: '{{BENEPBNSTC}}', 33: '{{BENEPBDEF}}',
    // SIT - 2026 Benefits
    35: '{{BENE26TTC}}', 36: '{{BENE26ETC}}', 37: '{{BENE26EP}}%', 38: '{{BENE26PTC}}',
    39: '{{BENE26PP}}%', 40: '{{BENE26FTC}}', 41: '{{BENE26FP}}%', 42: '{{BENE26IPTC}}',
    43: '{{BENE26BTC}}', 44: '{{BENE26NSTC}}', 45: '{{BENE26DEF}}',
    // Per-plan table (Target Completion Date cells at 60/70/80/90/100 left as literal dates)
    59: '{{BENESIHMTTC}}', 61: '{{BENESIHMETC}}', 62: '{{BENESIHMPTC}}', 63: '{{BENESIHMFTC}}',
    64: '{{BENESIHMIPTC}}', 65: '{{BENESIHMBTC}}', 66: '{{BENESIHMNSTC}}', 67: '{{BENESIHMDEFOPEN}}',
    69: '{{BENEACPTTC}}', 71: '{{BENEACPETC}}', 72: '{{BENEACPPTC}}', 73: '{{BENEACPFTC}}',
    74: '{{BENEACPIPTC}}', 75: '{{BENEACPBTC}}', 76: '{{BENEACPNSTC}}', 77: '{{BENEACPDEFOPEN}}',
    79: '{{BENEPRPTTC}}', 81: '{{BENEPRPETC}}', 82: '{{BENEPRPPTC}}', 83: '{{BENEPRPFTC}}',
    84: '{{BENEPRPIPTC}}', 85: '{{BENEPRPBTC}}', 86: '{{BENEPRPNSTC}}', 87: '{{BENEPRPDEFOPEN}}',
    89: '{{BENEMMEGWPTTC}}', 91: '{{BENEMMEGWPETC}}', 92: '{{BENEMMEGWPPTC}}', 93: '{{BENEMMEGWPFTC}}',
    94: '{{BENEMMEGWPIPTC}}', 95: '{{BENEMMEGWPBTC}}', 96: '{{BENEMMEGWPNSTC}}', 97: '{{BENEMMEGWPDEFOPEN}}',
    99: '{{BENENEOHTTC}}', 101: '{{BENENEOHETC}}', 102: '{{BENENEOHPTC}}', 103: '{{BENENEOHFTC}}',
    104: '{{BENENEOHIPTC}}', 105: '{{BENENEOHBTC}}', 106: '{{BENENEOHNSTC}}', 107: '{{BENENEOHDEFOPEN}}',
    // Grand Total
    109: '{{ACTIVEBENTTC}}', 111: '{{ACTIVEBENETC}}', 112: '{{ACTIVEBENPTC}}', 113: '{{ACTIVEBENFTC}}',
    114: '{{ACTIVEBENIPTC}}', 115: '{{ACTIVEBENBTC}}', 116: '{{ACTIVEBENNSTC}}', 117: '{{ACTIVEBENB}}',
    134: '{{AI_BENEFITS_STATUS}}',
  },

  'ppt/slides/slide6.xml': {
    2: 'As-of: {{Date}}',
    10: '{{BENESIHMDEFTOTAL}}', 11: '{{BENESIHMDEFOPEN}}', 12: '{{BENESIHMDEFCLOSED}}', 13: '{{BENESIHMDEFRESOLVED}}',
    15: '{{BENEACPDEFTOTAL}}', 16: '{{BENEACPDEFOPEN}}', 17: '{{BENEACPDEFCLOSED}}', 18: '{{BENEACPDEFRESOLVED}}',
    20: '{{BENEPRPDEFTOTAL}}', 21: '{{BENEPRPDEFOPEN}}', 22: '{{BENEPRPDEFCLOSED}}', 23: '{{BENEPRPDEFRESOLVED}}',
    25: '{{BENENEOHDEFTOTAL}}', 26: '{{BENENEOHDEFOPEN}}', 27: '{{BENENEOHDEFCLOSED}}', 28: '{{BENENEOHDEFRESOLVED}}',
    30: '{{BENEMMEGWPDEFTOTAL}}', 31: '{{BENEMMEGWPDEFOPEN}}', 32: '{{BENEMMEGWPDEFCLOSED}}', 33: '{{BENEMMEGWPDEFRESOLVED}}',
    35: '{{BENEDEFTOTAL}}', 36: '{{ACTIVEBENB}}', 37: '{{BENEDEFCLOSEDTOTAL}}', 38: '{{BENEDEFRESOLVEDTOTAL}}',
    45: '{{AI_BENEFITS_DEFECT_STATUS}}', 46: '', 47: '', 48: '', 49: '',
    52: '{{AI_BENEFITS_DEFECT_INSIGHTS}}', 53: '', 54: '', 55: '', 56: '', 57: '',
    60: '{{AI_BENEFITS_BURNDOWN_INSIGHTS}}', 61: '', 62: '', 63: '', 64: '', 65: '', 66: '',
  },

  'ppt/slides/slide7.xml': {
    6: 'As-of: {{Date}}',
    23: '{{ENROTTC}}', 24: '{{ENROETC}}', 25: '{{ENROEP}}%', 26: '{{ENROPTC}}',
    27: '{{ENROPP}}%', 28: '{{ENROFTC}}', 29: '{{ENROFP}}%', 30: '{{ENROIPTC}}',
    31: '{{ENROBTC}}', 32: '{{ENRONSTC}}', 33: '{{ENROB}}',
    47: '{{ENROTTC}}', 49: '{{ENROETC}}', 50: '{{ENROPTC}}', 51: '{{ENROFTC}}',
    52: '{{ENROIPTC}}', 53: '{{ENROBTC}}', 54: '{{ENRONSTC}}', 55: '{{ENROB}}',
    57: '{{ENROTTC}}', 59: '{{ENROETC}}', 60: '{{ENROPTC}}', 61: '{{ENROFTC}}',
    62: '{{ENROIPTC}}', 63: '{{ENROBTC}}', 64: '{{ENRONSTC}}', 65: '{{ENROB}}',
    81: '{{AI_ENROLLMENT_STATUS}}',
  },

  'ppt/slides/slide8.xml': {
    6: 'As-of: {{Date}}',
    23: '{{EDITTC}}', 24: '{{EDIETC}}', 25: '{{EDIEP}}%', 26: '{{EDIPTC}}',
    27: '{{EDIPP}}%', 28: '{{EDIFTC}}', 29: '{{EDIFP}}%', 30: '{{EDIIPTC}}',
    31: '{{EDIBTC}}', 32: '{{EDINSTC}}', 33: '{{EDIB}}',
    47: '{{EDITTC}}', 49: '{{EDIETC}}', 50: '{{EDIPTC}}', 51: '{{EDIFTC}}',
    52: '{{EDIIPTC}}', 53: '{{EDIBTC}}', 54: '{{EDINSTC}}', 55: '{{EDIB}}',
    57: '{{EDITTC}}', 59: '{{EDIETC}}', 60: '{{EDIPTC}}', 61: '{{EDIFTC}}',
    62: '{{EDIIPTC}}', 63: '{{EDIBTC}}', 64: '{{EDINSTC}}', 65: '{{EDIB}}',
    79: '{{AI_EDI_STATUS}}',
  },
};

function applyMap(xml, map) {
  let i = -1;
  return xml.replace(/<a:t>([^<]*)<\/a:t>/g, (full, _text) => {
    i++;
    if (!(i in map)) return full;
    // Escape XML special chars in the replacement (tokens/literals here are
    // plain ASCII/braces/%, so this is just defensive).
    const escaped = map[i].replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return `<a:t>${escaped}</a:t>`;
  });
}

const zip = new PizZip(fs.readFileSync(SOURCE, 'binary'));

for (const [slidePath, map] of Object.entries(SLIDE_MAPS)) {
  const file = zip.file(slidePath);
  if (!file) {
    console.warn(`⚠  ${slidePath} not found in source deck — skipping`);
    continue;
  }
  const xml = applyMap(file.asText(), map);
  zip.file(slidePath, xml);
  console.log(`✓ ${slidePath}: ${Object.keys(map).length} run(s) tokenized`);
}

fs.writeFileSync(OUTPUT, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
console.log(`\nTemplate written → ${OUTPUT}`);
