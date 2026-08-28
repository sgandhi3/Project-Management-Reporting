// Narrative Refresh extension — no external API, no API key required.
//
// The .pptx template has several free-text status sentences (executive summary
// bullets, per-workstream "Overall ... testing is On Track/At Risk ..." lines,
// and PDM defect-detail cells) that go stale week to week — hardcoded dates,
// counts, and plan names. This extension composes each one fresh from this
// week's actual ADO data and stores the results on `data._aiNarratives`, which
// variables.js exposes as {{AI_...}} tokens for extensions/ppt.js to substitute.
//
// Two sources, per key, in priority order:
//   1. AGENT-WRITTEN OVERRIDE — if ./ai-narrative-input.json exists (written by
//      whatever ran gather-data.js — e.g. a Claude Code agent that read the
//      --preview data and composed real sentences itself, no API key needed
//      since it's using its own session rather than calling out to a model),
//      its values are used verbatim, key by key.
//   2. RULE-BASED FALLBACK — for any key missing from the override (or when
//      there's no override file at all, e.g. a plain `npm run generate`),
//      computed here from plain thresholds/keyword matching. No names of
//      individuals are ever included — only counts and defect themes.
//
// The status thresholds and category keywords below are judgment calls —
// tune them for your program if "on track" / "at risk" / "off track" should
// mean something different for you.
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

function pct(n, d) { return d ? Math.round((n / d) * 100) : 0; }
function plural(n, word) { return `${n} ${word}${n === 1 ? '' : 's'}`; }

// ─── Status judgment ───────────────────────────────────────────────────────
// Judged on pass rate and blockers among what's been executed so far, NOT on
// how much of the plan has been executed — low volume early in a cycle isn't
// itself trouble. Tune these thresholds to match how your program defines
// the labels; this is a rough proxy for what used to be a human judgment call.
function classifyStatus(s) {
  if (!s || !s.executed) return 'not yet started';
  const passPct = pct(s.passed, s.executed);
  if ((s.blocked ?? 0) > 0 || passPct < 50) return 'off track';
  if (passPct < 85 || s.failed > 0) return 'at risk';
  return 'on track';
}

// ─── Defect helpers ──────────────────────────────────────────────────────────
// Deliberately no owner/assignee anywhere below — counts and themes only.

// Keyword-based categorization of defect titles — best-effort, not exact.
const CATEGORY_KEYWORDS = [
  ['data quality',  ['invalid', 'missing', 'incorrect', 'mismatch', 'duplicate', 'bad data']],
  ['data mapping',  ['mapping', 'map', 'roster', 'crosswalk', 'cross-walk']],
  ['configuration', ['config', 'setup', 'setting', 'benefit provision']],
  ['integration',   ['integration', 'interface', 'edifecs', 'edi ']],
];
function categorize(title) {
  const t = (title || '').toLowerCase();
  for (const [cat, kws] of CATEGORY_KEYWORDS) {
    if (kws.some(k => t.includes(k))) return cat;
  }
  return null; // no confident category — omit rather than say "other"
}

// Named-category breakdown only (drops uncategorized defects from the list —
// better to say nothing than a meaningless "(primarily other)").
function categoryBreakdown(bugs) {
  const counts = {};
  for (const b of bugs) {
    const cat = categorize(b.title);
    if (cat) counts[cat] = (counts[cat] || 0) + 1;
  }
  return Object.entries(counts).sort((a, b) => b[1] - a[1]);
}

function categoryClause(bugs) {
  const cats = categoryBreakdown(bugs);
  if (!cats.length) return '';
  return ` (${cats.map(([c, n]) => `${plural(n, 'defect')} related to ${c}`).join(', ')})`;
}

function defectDetailSentence(bugs, severityLabel) {
  if (!bugs.length) return `No open ${severityLabel} PDM defects at this time.`;
  return `${plural(bugs.length, 'open ' + severityLabel + ' PDM defect')}${categoryClause(bugs)}.`;
}

// ─── Sentence builders ────────────────────────────────────────────────────────

function buildOverallStatus(data) {
  const entries    = Object.entries(data.stats);
  const started    = entries.filter(([, s]) => s.executed > 0).sort((a, b) => b[1].executed - a[1].executed);
  const notStarted = entries.filter(([, s]) => !s.executed).map(([n]) => n);
  const c = data.consolidatedData;
  const status = classifyStatus(c);
  const overallPassPct = pct(c.passed, c.executed);

  let sentence = `Overall, SIT execution is ${status}, with ${c.executed} of ${c.planned} planned test cases executed`
    + (c.executed ? ` (${overallPassPct}% pass rate).` : '.');

  if (started.length) {
    sentence += ` ${started.map(([n]) => n).join(', ')} ${started.length > 1 ? 'lead' : 'leads'} execution progress`;
    sentence += notStarted.length
      ? `; ${notStarted.join(', ')} ${notStarted.length > 1 ? 'have' : 'has'} not yet begun executing test cases.`
      : '.';
  } else if (notStarted.length) {
    sentence += ' No workstream has begun executing test cases yet.';
  }
  return sentence;
}

function buildPdmIterationUpdate(data) {
  const s = data.stats.PDM;
  if (!s || !s.planned) return 'PDM execution data is not yet available.';
  return `PDM execution is underway, with ${s.executed} of ${s.planned} planned test cases executed to date (${pct(s.executed, s.planned)}%).`;
}

function buildPdmDefectSummary(data) {
  const bugs = data.bugs?.PDM ?? [];
  if (!bugs.length) return 'PDM currently has no open ADO defects.';
  const high = bugs.filter(b => ['1 - Critical', '2 - High'].includes(b.severity)).length;
  const med  = bugs.filter(b => b.severity === '3 - Medium').length;
  const low  = bugs.filter(b => b.severity === '4 - Low').length;
  return `PDM has ${plural(bugs.length, 'open ADO defect')} (${high} critical/high, ${med} medium, ${low} low severity).`;
}

function buildBenefitsUpdate(data) {
  const s = data.stats.Benefits;
  const bugs = data.bugs?.Benefits ?? [];
  if (!s || !s.executed) return `Priority Benefits testing has ${plural(bugs.length, 'open defect')}; execution has not yet begun.`;
  return `Priority Benefits testing has executed ${s.executed} of ${s.planned} test cases (${pct(s.passed, s.executed)}% pass rate), with ${plural(bugs.length, 'open defect')} being worked.`;
}

function buildWorkstreamStatusLine(label, s, bugs) {
  const status = classifyStatus(s);
  return `Overall ${label} testing is ${status}, with ${plural(bugs.length, 'open defect')}${categoryClause(bugs)}.`;
}

function computeFallback(data) {
  const pdmBugs   = data.bugs?.PDM ?? [];
  const pdmHigh   = pdmBugs.filter(b => b.severity === '1 - Critical' || b.severity === '2 - High');
  const pdmMedLow = pdmBugs.filter(b => b.severity === '3 - Medium' || b.severity === '4 - Low');

  // Template has two lines for the critical/high-severity detail cell — line 1
  // states the count/category breakdown, line 2 gives one illustrative example
  // (defect title only — never a person's name).
  const exampleTitle = pdmHigh[0]?.title;

  return {
    overallStatus:         buildOverallStatus(data),
    pdmIterationUpdate:    buildPdmIterationUpdate(data),
    pdmDefectSummary:      buildPdmDefectSummary(data),
    benefitsUpdate:        buildBenefitsUpdate(data),
    pdmCursoryStatus:      buildWorkstreamStatusLine('PDM Cursory Review', data.stats.PDM, pdmBugs),
    benefitsStatus:        buildWorkstreamStatusLine('Benefits SIT', data.stats.Benefits, data.bugs?.Benefits ?? []),
    enrollmentStatus:      buildWorkstreamStatusLine('Enrollment SIT', data.stats.Enrollment, data.bugs?.Enrollment ?? []),
    ediStatus:             buildWorkstreamStatusLine('EDI SIT', data.stats.EDI, data.bugs?.EDI ?? []),
    pdmDataQualityDetail1: pdmHigh.length
      ? `${plural(pdmHigh.length, 'critical/high-severity PDM defect')} open${categoryClause(pdmHigh)}.`
      : 'No open critical/high-severity PDM defects at this time.',
    pdmDataQualityDetail2: exampleTitle ? `Example: "${exampleTitle}".` : '',
    pdmDataMappingDetail:  defectDetailSentence(pdmMedLow, 'medium/low-severity'),
  };
}

function loadOverride() {
  if (!fs.existsSync(OVERRIDE_PATH)) return {};
  try {
    const parsed = JSON.parse(fs.readFileSync(OVERRIDE_PATH, 'utf8'));
    console.log(`  Using agent-written sentences from ${OVERRIDE_PATH}`);
    return parsed;
  } catch (e) {
    console.warn(`  ⚠  Could not parse ${OVERRIDE_PATH} — falling back to rule-based sentences (${e.message})`);
    return {};
  }
}

export async function generate(data) {
  data._aiNarratives = data._aiNarratives || {};

  console.log('\nRefreshing status sentences from current data...');

  const override = loadOverride();
  const fallback = computeFallback(data);

  for (const key of NARRATIVE_KEYS) {
    const overrideVal = override[key];
    data._aiNarratives[key] = (typeof overrideVal === 'string' && overrideVal.trim())
      ? overrideVal.trim()
      : fallback[key];
  }

  console.log('Refreshed sentences:');
  for (const key of NARRATIVE_KEYS) {
    const v = data._aiNarratives[key];
    console.log(`  ${key}: ${v ? v.slice(0, 90) + (v.length > 90 ? '…' : '') : '(empty)'}`);
  }
}
