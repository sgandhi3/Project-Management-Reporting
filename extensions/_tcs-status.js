// Colors the "Status" cell on the Test Case Planning Summary slide (one per
// workstream) based on actual-vs-planned pacing — see _tcs-planning.js for
// the shared red/yellow/green rule. The cell has no visible text (its
// {{<WS>TCSSTATUS}} token always resolves to '' in variables.js); this
// extension locates it by that same token string BEFORE the normal
// {{TOKEN}} replacement pass blanks it, and swaps the template's neutral
// placeholder fill color for the computed one.
//
// Runs in the same early phase as _dynamic-benefits.js / _burndown-charts.js
// (see extensions/ppt.js) — i.e. before the generic text-token replacement.

import { computeTcsPlanning } from './_tcs-planning.js';

const SLIDE = 'ppt/slides/slide11.xml';
const PLACEHOLDER_FILL = 'BDBDBD';

const TARGETS = [
  { token: 'PDMTCSSTATUS',  ws: 'PDM' },
  { token: 'BENETCSSTATUS', ws: 'Benefits' },
  { token: 'ENROTCSSTATUS', ws: 'Enrollment' },
  { token: 'EDITCSSTATUS',  ws: 'EDI' },
];

// Benefits-only active-plan filter — mirrors variables.js's activeBenSum so
// this extension's numbers can never drift from what the table itself shows.
function activeBenSum(data, field) {
  return Object.entries(data.subStats?.Benefits ?? {})
    .filter(([key, s]) => key.split(' / ').length === 2 && s.executed > 0)
    .reduce((acc, [, s]) => acc + (s[field] ?? 0), 0);
}

function adoTotal(data, ws) {
  if (ws === 'PDM') return data.subStats?.PDM?.SIT?.planned ?? 0; // SIT only, no Cursory
  return data.stats?.[ws]?.planned ?? 0;
}
function adoExecuted(data, ws) {
  if (ws === 'PDM') return data.subStats?.PDM?.SIT?.executed ?? 0;
  return data.stats?.[ws]?.executed ?? 0;
}

// a:tc never nests within a:tc — safe to match top-level occurrences.
function matchCells(xml) {
  const re = /<a:tc\b[^>]*?(?:\/>|>[\s\S]*?<\/a:tc>)/g;
  const out = [];
  let m;
  while ((m = re.exec(xml))) out.push({ start: m.index, end: m.index + m[0].length, text: m[0] });
  return out;
}

export function apply(zip, data) {
  const file = zip.file(SLIDE);
  if (!file) return;
  let xml = file.asText();

  for (const { token, ws } of TARGETS) {
    const cell = matchCells(xml).find(c => c.text.includes(`{{${token}}}`));
    if (!cell) {
      console.warn(`  ⚠  _tcs-status: couldn't find cell for {{${token}}} — skipping`);
      continue;
    }
    const { color } = computeTcsPlanning(adoTotal(data, ws), adoExecuted(data, ws), ws);
    const newText = cell.text.replace(`<a:srgbClr val="${PLACEHOLDER_FILL}"/>`, `<a:srgbClr val="${color}"/>`);
    if (newText === cell.text) {
      console.warn(`  ⚠  _tcs-status: placeholder fill not found in {{${token}}}'s cell — skipping`);
      continue;
    }
    xml = xml.slice(0, cell.start) + newText + xml.slice(cell.end);
  }

  zip.file(SLIDE, xml);
}
