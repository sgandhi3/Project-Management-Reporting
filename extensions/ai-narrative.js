// Narrative Refresh extension — no external API, no API key required.
//
// The .pptx template has several free-text status sentences (executive summary
// bullets, per-workstream "Overall ... testing is On Track/At Risk ..." lines,
// and PDM defect-detail cells) that go stale week to week — hardcoded dates,
// counts, and plan names. This extension composes each one fresh from this
// week's actual ADO data using plain rules (status thresholds, defect counts,
// top owner, keyword-based categorization) and stores the results on
// `data._aiNarratives`, which variables.js exposes as {{AI_...}} tokens for
// extensions/ppt.js to substitute.
//
// The status thresholds and category keywords below are judgment calls —
// tune them for your program if "on track" / "at risk" / "off track" should
// mean something different for you.

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

function topOwner(bugs) {
  const counts = {};
  for (const b of bugs) counts[b.owner] = (counts[b.owner] || 0) + 1;
  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  return sorted[0]?.[0] || null;
}

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
  return 'other';
}

function categoryBreakdown(bugs) {
  const counts = {};
  for (const b of bugs) counts[categorize(b.title)] = (counts[categorize(b.title)] || 0) + 1;
  return Object.entries(counts).sort((a, b) => b[1] - a[1]);
}

function defectDetailSentence(bugs, severityLabel) {
  if (!bugs.length) return `No open ${severityLabel} PDM defects at this time.`;
  const cats  = categoryBreakdown(bugs);
  const owner = topOwner(bugs);
  const catText = cats.map(([c, n]) => `${plural(n, 'defect')} related to ${c}`).join(', ');
  return `${plural(bugs.length, 'open ' + severityLabel + ' PDM defect')} (${catText})`
    + (owner ? `, the largest share owned by ${owner}.` : '.');
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
  const owner  = topOwner(bugs);
  let sentence = `Overall ${label} testing is ${status}, with ${plural(bugs.length, 'open defect')}`;
  if (bugs.length) {
    const cats = categoryBreakdown(bugs);
    sentence += ` (primarily ${cats[0][0]})`;
    if (owner) sentence += `, largest share owned by ${owner}`;
  }
  return sentence + '.';
}

export async function generate(data) {
  data._aiNarratives = data._aiNarratives || {};

  console.log('\nRefreshing status sentences from current data...');

  const pdmBugs = data.bugs?.PDM ?? [];
  const pdmHigh   = pdmBugs.filter(b => b.severity === '1 - Critical' || b.severity === '2 - High');
  const pdmMedLow = pdmBugs.filter(b => b.severity === '3 - Medium' || b.severity === '4 - Low');

  // The template has two lines for the critical/high-severity detail cell —
  // line 1 states the count/category breakdown, line 2 names the top owner.
  const dqCats  = categoryBreakdown(pdmHigh);
  const dqOwner = topOwner(pdmHigh);

  data._aiNarratives.overallStatus         = buildOverallStatus(data);
  data._aiNarratives.pdmIterationUpdate    = buildPdmIterationUpdate(data);
  data._aiNarratives.pdmDefectSummary      = buildPdmDefectSummary(data);
  data._aiNarratives.benefitsUpdate        = buildBenefitsUpdate(data);
  data._aiNarratives.pdmCursoryStatus      = buildWorkstreamStatusLine('PDM Cursory Review', data.stats.PDM, pdmBugs);
  data._aiNarratives.benefitsStatus        = buildWorkstreamStatusLine('Benefits SIT', data.stats.Benefits, data.bugs?.Benefits ?? []);
  data._aiNarratives.enrollmentStatus      = buildWorkstreamStatusLine('Enrollment SIT', data.stats.Enrollment, data.bugs?.Enrollment ?? []);
  data._aiNarratives.ediStatus             = buildWorkstreamStatusLine('EDI SIT', data.stats.EDI, data.bugs?.EDI ?? []);
  data._aiNarratives.pdmDataQualityDetail1 = pdmHigh.length
    ? `${plural(pdmHigh.length, 'critical/high-severity PDM defect')} open (${dqCats.map(([c, n]) => `${plural(n, 'defect')} related to ${c}`).join(', ')}).`
    : 'No open critical/high-severity PDM defects at this time.';
  data._aiNarratives.pdmDataQualityDetail2 = (pdmHigh.length && dqOwner) ? `Largest share owned by ${dqOwner}.` : '';
  data._aiNarratives.pdmDataMappingDetail  = defectDetailSentence(pdmMedLow, 'medium/low-severity');

  console.log('Refreshed sentences:');
  for (const key of NARRATIVE_KEYS) {
    const v = data._aiNarratives[key];
    console.log(`  ${key}: ${v ? v.slice(0, 90) + (v.length > 90 ? '…' : '') : '(empty)'}`);
  }
}
