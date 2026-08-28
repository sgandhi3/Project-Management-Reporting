// Narrative Refresh extension — no rule-based generation, no external API.
//
// The .pptx template has several free-text status sentences (executive summary
// bullets, per-workstream "Overall ... testing is On Track/At Risk ..." lines,
// and PDM defect-detail cells) that go stale week to week. Their tokens
// ({{AI_...}}) are only ever filled from an AGENT-WRITTEN OVERRIDE file —
// ai-narrative-input.json in the project root, written by whatever ran
// gather-data.js (e.g. a Claude Code agent that read --narrative-data and
// composed real sentences itself, no API key needed since it's using its own
// session rather than calling out to a model).
//
// If that file doesn't exist, or a key is missing from it, that token is left
// blank — there is no generated fallback text.
import fs   from 'fs';
import path from 'path';

const OVERRIDE_PATH = path.join(process.cwd(), 'ai-narrative-input.json');

const NARRATIVE_KEYS = [
  'overallStatus',
  'pdmIterationUpdate',
  'pdmDefectSummary',
  'benefitsUpdate',
  'pdmCursoryStatus',
  'benefitsStatus',
  'enrollmentStatus',
  'ediStatus',
  'pdmDataQualityDetail1',
  'pdmDataQualityDetail2',
  'pdmDataMappingDetail',
];

function loadOverride() {
  if (!fs.existsSync(OVERRIDE_PATH)) return {};
  try {
    const parsed = JSON.parse(fs.readFileSync(OVERRIDE_PATH, 'utf8'));
    console.log(`  Using agent-written sentences from ${OVERRIDE_PATH}`);
    return parsed;
  } catch (e) {
    console.warn(`  ⚠  Could not parse ${OVERRIDE_PATH} — leaving AI_* tokens blank (${e.message})`);
    return {};
  }
}

export async function generate(data) {
  data._aiNarratives = data._aiNarratives || {};

  const override = loadOverride();
  for (const key of NARRATIVE_KEYS) {
    const val = override[key];
    data._aiNarratives[key] = typeof val === 'string' ? val.trim() : '';
  }

  const filled = NARRATIVE_KEYS.filter(k => data._aiNarratives[k]).length;
  console.log(`\nNarrative sentences: ${filled}/${NARRATIVE_KEYS.length} filled`
    + (filled ? '' : ' — no override file found, leaving AI_* tokens blank'));
}
