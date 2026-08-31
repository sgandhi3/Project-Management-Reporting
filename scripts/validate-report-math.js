// Cross-checks every number actually printed in a generated report against
// every other number it should agree with — catches template/token-wiring
// bugs (e.g. a token pointed at the wrong data) that a data-layer check
// can't see, since it reads the final rendered PPTX text, not the raw fetch.
//
// Checks performed:
//   1. Each detail slide's sub-suite rows sum to that table's own Grand
//      Total row (per column).
//   2. Each detail slide's "SIT" row matches the corresponding workstream
//      row on the executive summary slide.
//   3. The executive summary's Grand Total row equals the sum of the three
//      workstream rows.
//   4. The executive summary's stat boxes match the Grand Total row.
//   5. Every displayed percentage matches what its underlying counts
//      recompute to (±1 point tolerance for independent rounding).
//
// Columns are identified by their HEADER LABEL, not fixed position — the
// template has already had a column removed (Target Completion Date) and a
// column merged (Defects, vertically spanned across sub-suite rows) once
// already, which silently broke a position-based version of this script.
// Reading columns by label survives future column add/remove/reorder edits
// as long as the header text stays recognizable.
//
// Usage: node scripts/validate-report-math.js <path-to-generated-report.pptx>
import fs from 'fs';
import PizZip from 'pizzip';

const REPORT = process.argv[2];
if (!REPORT || !fs.existsSync(REPORT)) {
  console.error('Usage: node scripts/validate-report-math.js <path-to-report.pptx>');
  process.exit(1);
}

const zip = new PizZip(fs.readFileSync(REPORT, 'binary'));
const failures = [];
const note = (msg) => failures.push(msg);
const num = (s) => {
  const cleaned = String(s ?? '').replace(/[%*]/g, '').trim();
  if (cleaned === '' || cleaned === '--') return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
};

function slideXml(n) {
  const f = zip.file(`ppt/slides/slide${n}.xml`);
  return f ? f.asText() : null;
}

function extractTables(xml) {
  const tables = [];
  for (const tblMatch of xml.matchAll(/<a:tbl>[\s\S]*?<\/a:tbl>/g)) {
    const rows = [];
    for (const trMatch of tblMatch[0].matchAll(/<a:tr[ >][\s\S]*?<\/a:tr>/g)) {
      const cells = [];
      for (const tcMatch of trMatch[0].matchAll(/<a:tc[ >][\s\S]*?<\/a:tc>/g)) {
        const texts = [...tcMatch[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(m => m[1]);
        cells.push(texts.join(''));
      }
      rows.push(cells);
    }
    tables.push(rows);
  }
  return tables;
}

function allShapes(xml) {
  const shapes = [];
  for (const m of xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const cnv = m[0].match(/<p:cNvPr id="(\d+)" name="([^"]*)"/);
    if (!cnv) continue;
    const text = [...m[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(x => x[1]).join('');
    shapes.push({ id: cnv[1], name: cnv[2], text });
  }
  return shapes;
}

// Maps a header row's cell text to field names, by column CONTENT rather than
// position. "Executed" and "Execution %" both start with "Exec" — percentage
// columns are disambiguated first by requiring a literal "%" in the header.
function buildColumnMap(headerRow) {
  const map = {};
  headerRow.forEach((cell, idx) => {
    const t = (cell || '').trim();
    if (idx === 0) return; // first column is always the row label (Phase/Workstream/Benefit Type/Workstreams)
    if (/%/.test(t)) {
      if (/^exec/i.test(t)) map.execPct = idx;
      else if (/^pass/i.test(t)) map.passPct = idx;
      else if (/^fail/i.test(t)) map.failPct = idx;
      return;
    }
    if (/^planned/i.test(t)) map.planned = idx;
    else if (/^executed$/i.test(t)) map.executed = idx;
    else if (/^passed$/i.test(t)) map.passed = idx;
    else if (/^failed$/i.test(t)) map.failed = idx;
    else if (/in\s*progress/i.test(t)) map.inProgress = idx;
    else if (/^blocked$/i.test(t)) map.blocked = idx;
    else if (/not\s*started/i.test(t)) map.notStarted = idx;
    else if (/defects/i.test(t)) map.defects = idx;
    // anything else (e.g. "Target Completion Date") is ignored, not required
  });
  return map;
}

function rowValues(row, colMap) {
  const out = {};
  for (const field of ['planned', 'executed', 'execPct', 'passed', 'passPct', 'failed', 'failPct', 'inProgress', 'blocked', 'notStarted', 'defects']) {
    out[field] = colMap[field] !== undefined ? num(row[colMap[field]]) : null;
  }
  return out;
}

// Stat boxes are identified by their paired label text (e.g. "Total Test
// Cases", "Passed", "Executed") rather than hardcoded shape IDs, since the
// template's shape IDs shift whenever the stat-box layout is edited — the
// label text is the one thing guaranteed to still mean the same thing.
const STAT_BOX_LABELS = [
  [/^Total Test Cases/i, 'TTC'],
  [/^Passed/i, 'PTC'],
  [/^Failed/i, 'FTC'],
  [/^Open Defects/i, 'TB'],
  [/^Executed/i, 'ETC'],
  [/^In Progress/i, 'IPTC'],
];

function extractStatBoxes(xml) {
  const shapes = allShapes(xml);
  const boxes = { TTC: null, PTC: null, FTC: null, TB: null, ETC: null, IPTC: null };
  for (let i = 0; i < shapes.length; i++) {
    if (!/Num$/.test(shapes[i].name)) continue;
    let label = null;
    for (let j = i + 1; j < shapes.length; j++) {
      if (/Num$/.test(shapes[j].name)) break; // next Num with no Lbl in between — unpaired
      if (/Lbl$/.test(shapes[j].name)) { label = shapes[j].text; break; }
    }
    if (label == null) continue;
    const match = STAT_BOX_LABELS.find(([re]) => re.test(label.trim()));
    if (match) boxes[match[1]] = num(shapes[i].text);
  }
  return boxes;
}

// ── Slide 2: executive summary ──────────────────────────────────────────────
const s2 = slideXml(2);
const s2Tables = extractTables(s2);
const s2ColMap = buildColumnMap(s2Tables[0][1]); // header row is index 1
const execRows = {};
for (const row of s2Tables[0]) {
  const label = row[0];
  if (['PDM', 'Enrollment', 'EDI', 'Grand Total'].includes(label)) {
    execRows[label] = rowValues(row, s2ColMap);
  }
}

const statBoxes = extractStatBoxes(s2);

// ── Slides 3–5: detail slides ────────────────────────────────────────────────
const DETAIL_SLIDES = { 3: 'PDM', 4: 'EDI', 5: 'Enrollment' };
const detailTotals = {};

for (const [slideNum, ws] of Object.entries(DETAIL_SLIDES)) {
  const xml = slideXml(slideNum);
  const tables = extractTables(xml);

  // Table 0: Overall Testing Execution Summary — single "SIT" row
  const t0ColMap = buildColumnMap(tables[0][1]);
  const sitRow = tables[0].find(r => r[0] === 'SIT');
  const sit = rowValues(sitRow, t0ColMap);
  detailTotals[ws] = sit;

  // Table 1: sub-suite breakdown — sum rows should equal the Grand Total row
  const t1ColMap = buildColumnMap(tables[1][1]);
  const subRows = tables[1].filter(r => r[0] && r[0] !== 'Grand Total' && !/^(Workstream|Benefit Type)$/.test(r[0]) && num(r[t1ColMap.planned]) !== null);
  const grandRow = tables[1].find(r => r[0] === 'Grand Total');
  const grand = rowValues(grandRow, t1ColMap);

  const sumFields = ['planned', 'executed', 'passed', 'failed', 'inProgress', 'blocked', 'notStarted'];
  const summed = Object.fromEntries(sumFields.map(f => [f, 0]));
  for (const r of subRows) {
    const v = rowValues(r, t1ColMap);
    for (const f of sumFields) summed[f] += v[f] || 0;
    // defects intentionally excluded from the sum check — sub-suite defect counts are known to be
    // hardcoded to 0 or vertically merged into one workstream-wide value (ADO has no per-sub-suite
    // area path), so a per-row sum wouldn't mean anything meaningful to compare against Grand Total.
  }
  for (const field of sumFields) {
    if (summed[field] !== grand[field]) {
      note(`Slide ${slideNum} (${ws}): sub-suite rows sum to ${field}=${summed[field]} but Grand Total row shows ${grand[field]}`);
    }
  }

  // Table1 Grand Total row should equal Table0's SIT row (same workstream, same slide)
  for (const field of sumFields) {
    if (grand[field] !== sit[field]) {
      note(`Slide ${slideNum} (${ws}): detail table's Grand Total ${field}=${grand[field]} doesn't match Overall Summary SIT row ${field}=${sit[field]}`);
    }
  }

  // Detail slide's SIT row should match slide 2's executive-summary row for this workstream
  const execRow = execRows[ws];
  if (!execRow) {
    note(`Slide 2: no executive-summary row found for workstream "${ws}"`);
  } else {
    for (const field of ['planned', 'executed', 'passed', 'failed', 'inProgress', 'defects']) {
      if (execRow[field] !== sit[field]) {
        note(`Slide 2's ${ws} row (${field}=${execRow[field]}) doesn't match Slide ${slideNum}'s SIT row (${field}=${sit[field]})`);
      }
    }
  }
}

// ── Grand Total row = sum of PDM + Enrollment + EDI ─────────────────────────
const gt = execRows['Grand Total'];
if (gt) {
  const sumFields = ['planned', 'executed', 'passed', 'failed', 'inProgress', 'defects'];
  const summedGT = Object.fromEntries(sumFields.map(f => [f, 0]));
  for (const ws of ['PDM', 'Enrollment', 'EDI']) {
    if (!execRows[ws]) continue;
    for (const f of sumFields) summedGT[f] += execRows[ws][f] ?? 0;
  }
  for (const f of sumFields) {
    if (summedGT[f] !== gt[f]) {
      note(`Slide 2 Grand Total ${f}=${gt[f]} but PDM+Enrollment+EDI sums to ${summedGT[f]}`);
    }
  }

  // Stat boxes should match the Grand Total row
  const statChecks = [['TTC', 'planned'], ['ETC', 'executed'], ['PTC', 'passed'], ['FTC', 'failed'], ['IPTC', 'inProgress'], ['TB', 'defects']];
  for (const [box, field] of statChecks) {
    if (statBoxes[box] !== gt[field]) {
      note(`Slide 2 stat box ${box}=${statBoxes[box]} doesn't match Grand Total row's ${field}=${gt[field]}`);
    }
  }
}

// ── Recompute every displayed percentage from its underlying counts ────────
function checkPct(label, displayed, numerator, denominator) {
  if (displayed === null || numerator === null || denominator === null) return;
  const expected = denominator ? Math.round((numerator / denominator) * 100) : 0;
  if (Math.abs(expected - displayed) > 1) {
    note(`${label}: displayed ${displayed}% but ${numerator}/${denominator} recomputes to ${expected}%`);
  }
}
for (const ws of ['PDM', 'Enrollment', 'EDI', 'Grand Total']) {
  const r = execRows[ws];
  if (!r) continue;
  checkPct(`Slide 2 ${ws} exec%`, r.execPct, r.executed, r.planned);
  checkPct(`Slide 2 ${ws} pass%`, r.passPct, r.passed, r.executed);
  checkPct(`Slide 2 ${ws} fail%`, r.failPct, r.failed, r.executed);
}
for (const [slideNum, ws] of Object.entries(DETAIL_SLIDES)) {
  const sit = detailTotals[ws];
  checkPct(`Slide ${slideNum} SIT exec%`, sit.execPct, sit.executed, sit.planned);
  checkPct(`Slide ${slideNum} SIT pass%`, sit.passPct, sit.passed, sit.executed);
  checkPct(`Slide ${slideNum} SIT fail%`, sit.failPct, sit.failed, sit.executed);
}

// ── Report ───────────────────────────────────────────────────────────────────
if (failures.length === 0) {
  console.log(`✓ Math validation passed — every cross-slide number is internally consistent (${REPORT})`);
  process.exit(0);
} else {
  console.error(`✗ Math validation FAILED — ${failures.length} inconsistency(ies) found in ${REPORT}:`);
  for (const f of failures) console.error(`  - ${f}`);
  process.exit(1);
}
