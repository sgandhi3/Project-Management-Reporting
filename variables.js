// This file controls what gets injected into each {{TOKEN}} in your PowerPoint template.
// Each entry is: TOKEN_NAME: d => <expression that reads from the data object>
//
// The data object `d` has this shape:
//   d.stats.<WorkstreamName>               — { planned, executed, passed, failed, notStarted, inProgress, blocked, paused }
//   d.subStats.<WorkstreamName>['path']    — same shape for sub-suites (breadcrumb from ADO suite tree)
//   d.<queryKey>.<WorkstreamName>          — array of items from that query (e.g. d.bugs.PDM)
//   d.consolidatedData                     — stats summed across all workstreams
//   d.<queryKey>By<Field>                  — group-by counts: { total, 'value1': count, ... }
//
// Run `node gather-data.js --preview` to see all available d.subStats keys.
// WorkstreamName must match the `name` field in the WORKSTREAMS array in config.js.

const pct  = (n, d) => d ? Math.round((n / d) * 100) : 0;
const sub  = (d, ws, path) => d.subStats?.[ws]?.[path] ?? {};
const stat = (d, ws, path, field) => sub(d, ws, path)[field] ?? 0;

// Sums a field across multiple sub-suite paths
const sumStat = (d, ws, paths, field) => paths.reduce((acc, p) => acc + stat(d, ws, p, field), 0);

// Benefits active-plan filter:
// Only count individual benefit plans (depth-2 keys like 'Priority / Signature HMO',
// 'HMO / Classic HMO NEOH') that have ≥1 execution.
// Plans with 0 executions are excluded; their planned cases don't count.
//
// MMO-specific — the "Priority" plans (Signature HMO, Access PPO, Premium
// PPO) already have their own template rows; PRIORITY_BENEFIT_KEYS lets
// extensions/_dynamic-benefits.js find any OTHER plan that's gone active
// and needs a row inserted for it. See that file for the row-insertion
// logic; this file only supplies the underlying data filters.
export const PRIORITY_BENEFIT_KEYS = new Set([
  'Priority / Signature HMO',
  'Priority / Access PPO (Premium PPO INN)',
  'Priority / Premium PPO',
]);

const activeBenEntries = d =>
  Object.entries(d.subStats?.Benefits ?? {})
    .filter(([key, s]) => key.split(' / ').length === 2 && s.executed > 0);

const activeBenSuites = d => activeBenEntries(d).map(([, s]) => s);

const activeBenSum = (d, field) =>
  activeBenSuites(d).reduce((acc, s) => acc + (s[field] ?? 0), 0);

// Active benefit plans NOT already covered by a hardcoded Priority row —
// these are what extensions/_dynamic-benefits.js inserts new rows for.
export const getExtraActiveBenefitPlans = d =>
  activeBenEntries(d)
    .filter(([key]) => !PRIORITY_BENEFIT_KEYS.has(key))
    .map(([key, s]) => ({ key, label: key.split(' / ').pop(), stats: s }));

// Grand total across all workstreams, using filtered Benefits
const activeGrandTotal = (d, field) =>
  Object.entries(d.stats)
    .filter(([ws]) => ws !== 'Benefits')
    .reduce((acc, [, s]) => acc + (s[field] ?? 0), 0)
  + activeBenSum(d, field);

// PDM Iteration 3 CPIMS has no parent key — sum the four individual sub-suites
const CPIMS3_PATHS = [
  'Cursory / Iteration 3 / Cpims-Prac',
  'Cursory / Iteration 3 / Cpims-Prac Role',
  'Cursory / Iteration 3 / CPIMS-ORG Location',
  'Cursory / Iteration 3 / CPIMS-ORG',
];

export const VARIABLE_MAP = {

  // ── Report Date ────────────────────────────────────────────────────────────
  Date: () => new Date().toLocaleDateString('en-US', {
    timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: 'numeric',
  }),

  // ── Grand Totals (Benefits filtered to active folders only) ──────────────
  // Excludes Benefits sub-suites with 0 executions (e.g. PPO, EGWP) from planned count.
  TTC:  d => activeGrandTotal(d, 'planned'),
  ETC:  d => activeGrandTotal(d, 'executed'),
  EP:   d => pct(activeGrandTotal(d, 'executed'),   activeGrandTotal(d, 'planned')),
  PTC:  d => activeGrandTotal(d, 'passed'),
  PP:   d => pct(activeGrandTotal(d, 'passed'),     activeGrandTotal(d, 'executed')),
  FTC:  d => activeGrandTotal(d, 'failed'),
  FP:   d => pct(activeGrandTotal(d, 'failed'),     activeGrandTotal(d, 'executed')),
  IPTC: d => activeGrandTotal(d, 'inProgress'),
  NSTC: d => activeGrandTotal(d, 'notStarted'),
  TB:   d => d.bugsBySeverity.total,

  // ── PDM — Cursory sub-suite (key: 'Cursory') ──────────────────────────────
  PDMCTTC:   d => stat(d, 'PDM', 'Cursory', 'planned'),
  PDMCETC:   d => stat(d, 'PDM', 'Cursory', 'executed'),
  PDMCEP:    d => pct(stat(d, 'PDM', 'Cursory', 'executed'),  stat(d, 'PDM', 'Cursory', 'planned')),
  PDMCPTC:   d => stat(d, 'PDM', 'Cursory', 'passed'),
  PDMCPP:    d => pct(stat(d, 'PDM', 'Cursory', 'passed'),    stat(d, 'PDM', 'Cursory', 'executed')),
  PDMCFTC:   d => stat(d, 'PDM', 'Cursory', 'failed'),
  PDMCFP:    d => pct(stat(d, 'PDM', 'Cursory', 'failed'),    stat(d, 'PDM', 'Cursory', 'executed')),
  PDMCIPTC:  d => stat(d, 'PDM', 'Cursory', 'inProgress'),
  PDMCB:     d => (d.bugs?.PDM ?? []).length,

  // ── PDM — SIT sub-suite (key: 'SIT') ──────────────────────────────────────
  PDMSITTC:   d => stat(d, 'PDM', 'SIT', 'planned'),
  PDMSITETC:  d => stat(d, 'PDM', 'SIT', 'executed'),
  PDMSITEP:   d => pct(stat(d, 'PDM', 'SIT', 'executed'),  stat(d, 'PDM', 'SIT', 'planned')),
  PDMSITPTC:  d => stat(d, 'PDM', 'SIT', 'passed'),
  PDMSITPP:   d => pct(stat(d, 'PDM', 'SIT', 'passed'),    stat(d, 'PDM', 'SIT', 'executed')),
  PDMSITFTC:  d => stat(d, 'PDM', 'SIT', 'failed'),
  PDMSITFP:   d => pct(stat(d, 'PDM', 'SIT', 'failed'),    stat(d, 'PDM', 'SIT', 'executed')),
  PDMSITIPTC: d => stat(d, 'PDM', 'SIT', 'inProgress'),
  PDMSITB:    d => 0,

  // ── PDM — Cursory Iteration breakdown ─────────────────────────────────────
  // ADO paths: 'Cursory / Iteration 2 / Iteration 2.0', 'Cursory / Iteration 2 / Iteration 2.1', 'Cursory / Iteration 3'
  PDMIT2TTC:   d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'planned'),
  PDMIT2ETC:   d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'executed'),
  PDMIT2EP:    d => pct(stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'executed'), stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'planned')),
  PDMIT2PTC:   d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'passed'),
  PDMIT2PP:    d => pct(stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'passed'),   stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'executed')),
  PDMIT2FTC:   d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'failed'),
  PDMIT2FP:    d => pct(stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'failed'),   stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'executed')),
  PDMIT2IPTC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'inProgress'),
  PDMIT2BTC:   d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'blocked'),
  PDMIT2NSTC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'notStarted'),
  PDMIT2B:     d => 0,

  PDMIT21TTC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'planned'),
  PDMIT21ETC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'executed'),
  PDMIT21EP:   d => pct(stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'executed'), stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'planned')),
  PDMIT21PTC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'passed'),
  PDMIT21PP:   d => pct(stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'passed'),   stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'executed')),
  PDMIT21FTC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'failed'),
  PDMIT21FP:   d => pct(stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'failed'),   stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'executed')),
  PDMIT21IPTC: d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'inProgress'),
  PDMIT21BTC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'blocked'),
  PDMIT21NSTC: d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'notStarted'),
  PDMIT21B:    d => 0,

  PDMIT3TTC:   d => stat(d, 'PDM', 'Cursory / Iteration 3', 'planned'),
  PDMIT3ETC:   d => stat(d, 'PDM', 'Cursory / Iteration 3', 'executed'),
  PDMIT3EP:    d => pct(stat(d, 'PDM', 'Cursory / Iteration 3', 'executed'), stat(d, 'PDM', 'Cursory / Iteration 3', 'planned')),
  PDMIT3PTC:   d => stat(d, 'PDM', 'Cursory / Iteration 3', 'passed'),
  PDMIT3PP:    d => pct(stat(d, 'PDM', 'Cursory / Iteration 3', 'passed'),   stat(d, 'PDM', 'Cursory / Iteration 3', 'executed')),
  PDMIT3FTC:   d => stat(d, 'PDM', 'Cursory / Iteration 3', 'failed'),
  PDMIT3FP:    d => pct(stat(d, 'PDM', 'Cursory / Iteration 3', 'failed'),   stat(d, 'PDM', 'Cursory / Iteration 3', 'executed')),
  PDMIT3IPTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3', 'inProgress'),
  PDMIT3BTC:   d => stat(d, 'PDM', 'Cursory / Iteration 3', 'blocked'),
  PDMIT3NSTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3', 'notStarted'),
  PDMIT3B:     d => 0,

  // ── PDM — Iteration 3 Data Source breakdown ────────────────────────────────
  // CPIMS has no parent suite — summed across 4 individual sub-suites
  PDMIT3CPTTC:  d => sumStat(d, 'PDM', CPIMS3_PATHS, 'planned'),
  PDMIT3CPETC:  d => sumStat(d, 'PDM', CPIMS3_PATHS, 'executed'),
  PDMIT3CPPTC:  d => sumStat(d, 'PDM', CPIMS3_PATHS, 'passed'),
  PDMIT3CPFTC:  d => sumStat(d, 'PDM', CPIMS3_PATHS, 'failed'),
  PDMIT3CPIPTC: d => sumStat(d, 'PDM', CPIMS3_PATHS, 'inProgress'),
  PDMIT3CPBTC:  d => sumStat(d, 'PDM', CPIMS3_PATHS, 'blocked'),
  PDMIT3CPNSTC: d => sumStat(d, 'PDM', CPIMS3_PATHS, 'notStarted'),
  PDMIT3CPB:    d => 0,

  PDMIT3ANTTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'planned'),
  PDMIT3ANETC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'executed'),
  PDMIT3ANPTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'passed'),
  PDMIT3ANFTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'failed'),
  PDMIT3ANIPTC: d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'inProgress'),
  PDMIT3ANBTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'blocked'),
  PDMIT3ANNSTC: d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'notStarted'),
  PDMIT3ANB:    d => 0,

  PDMIT3ROTTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'planned'),
  PDMIT3ROETC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'executed'),
  PDMIT3ROPTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'passed'),
  PDMIT3ROFTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'failed'),
  PDMIT3ROIPTC: d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'inProgress'),
  PDMIT3ROBTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'blocked'),
  PDMIT3RONSTC: d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'notStarted'),
  PDMIT3ROB:    d => 0,

  PDMIT3PNTTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'planned'),
  PDMIT3PNETC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'executed'),
  PDMIT3PNPTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'passed'),
  PDMIT3PNFTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'failed'),
  PDMIT3PNIPTC: d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'inProgress'),
  PDMIT3PNBTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'blocked'),
  PDMIT3PNNSTC: d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'notStarted'),
  PDMIT3PNB:    d => 0,

  // ── PDM — Defect counts by severity ───────────────────────────────────────
  PDMBCRIT: d => (d.bugs?.PDM ?? []).filter(b => b.severity === '1 - Critical').length,
  PDMBHIGH: d => (d.bugs?.PDM ?? []).filter(b => b.severity === '2 - High').length,
  PDMBMED:  d => (d.bugs?.PDM ?? []).filter(b => b.severity === '3 - Medium').length,
  PDMBLOW:  d => (d.bugs?.PDM ?? []).filter(b => b.severity === '4 - Low').length,

  // ── Benefits — Priority SIT (key: 'Priority') ─────────────────────────────
  BENEPBTTC:   d => stat(d, 'Benefits', 'Priority', 'planned'),
  BENEPBETC:   d => stat(d, 'Benefits', 'Priority', 'executed'),
  BENEPBEP:    d => pct(stat(d, 'Benefits', 'Priority', 'executed'), stat(d, 'Benefits', 'Priority', 'planned')),
  BENEPBPTC:   d => stat(d, 'Benefits', 'Priority', 'passed'),
  BENEPBPP:    d => pct(stat(d, 'Benefits', 'Priority', 'passed'),   stat(d, 'Benefits', 'Priority', 'executed')),
  BENEPBFTC:   d => stat(d, 'Benefits', 'Priority', 'failed'),
  BENEPBFP:    d => pct(stat(d, 'Benefits', 'Priority', 'failed'),   stat(d, 'Benefits', 'Priority', 'executed')),
  BENEPBIPTC:  d => stat(d, 'Benefits', 'Priority', 'inProgress'),
  BENEPBBTC:   d => stat(d, 'Benefits', 'Priority', 'blocked'),
  BENEPBNSTC:  d => stat(d, 'Benefits', 'Priority', 'notStarted'),
  BENEPBB:     d => (d.bugs?.Benefits ?? []).length,

  // ── Benefits — 2026 Benefits (active non-Priority folders only) ──────────
  // Uses activeBenSum so folders with 0 executions (PPO, EGWP) are excluded.
  BENE26TTC:   d => activeBenSum(d, 'planned')    - stat(d, 'Benefits', 'Priority', 'planned'),
  BENE26ETC:   d => activeBenSum(d, 'executed')   - stat(d, 'Benefits', 'Priority', 'executed'),
  BENE26EP:    d => pct(activeBenSum(d, 'executed') - stat(d, 'Benefits', 'Priority', 'executed'), activeBenSum(d, 'planned') - stat(d, 'Benefits', 'Priority', 'planned')),
  BENE26PTC:   d => activeBenSum(d, 'passed')     - stat(d, 'Benefits', 'Priority', 'passed'),
  BENE26PP:    d => pct(activeBenSum(d, 'passed')  - stat(d, 'Benefits', 'Priority', 'passed'),   activeBenSum(d, 'executed') - stat(d, 'Benefits', 'Priority', 'executed')),
  BENE26FTC:   d => activeBenSum(d, 'failed')     - stat(d, 'Benefits', 'Priority', 'failed'),
  BENE26FP:    d => pct(activeBenSum(d, 'failed')  - stat(d, 'Benefits', 'Priority', 'failed'),   activeBenSum(d, 'executed') - stat(d, 'Benefits', 'Priority', 'executed')),
  BENE26IPTC:  d => activeBenSum(d, 'inProgress') - stat(d, 'Benefits', 'Priority', 'inProgress'),
  BENE26BTC:   d => activeBenSum(d, 'blocked')    - stat(d, 'Benefits', 'Priority', 'blocked'),
  BENE26NSTC:  d => activeBenSum(d, 'notStarted') - stat(d, 'Benefits', 'Priority', 'notStarted'),
  BENE26B:     d => 0,

  // ── Benefits — active total (Priority + any active non-Priority plans) ───
  // Used by slide 4's own Grand Total row (BENEPBTTC alone under-counted
  // once a non-Priority plan went active — see activeBenSum above).
  ACTIVEBENTTC:  d => activeBenSum(d, 'planned'),
  ACTIVEBENETC:  d => activeBenSum(d, 'executed'),
  ACTIVEBENPTC:  d => activeBenSum(d, 'passed'),
  ACTIVEBENFTC:  d => activeBenSum(d, 'failed'),
  ACTIVEBENIPTC: d => activeBenSum(d, 'inProgress'),
  ACTIVEBENBTC:  d => activeBenSum(d, 'blocked'),
  ACTIVEBENNSTC: d => activeBenSum(d, 'notStarted'),
  ACTIVEBENB:    d => (d.bugs?.Benefits ?? []).length,

  // ── Benefits — Benefit type breakdown (Priority sub-suites) ───────────────
  BENESIHMTTC:  d => stat(d, 'Benefits', 'Priority / Signature HMO',               'planned'),
  BENESIHMETC:  d => stat(d, 'Benefits', 'Priority / Signature HMO',               'executed'),
  BENESIHMPTC:  d => stat(d, 'Benefits', 'Priority / Signature HMO',               'passed'),
  BENESIHMFTC:  d => stat(d, 'Benefits', 'Priority / Signature HMO',               'failed'),
  BENESIHMIPTC: d => stat(d, 'Benefits', 'Priority / Signature HMO',               'inProgress'),
  BENESIHMBTC:  d => stat(d, 'Benefits', 'Priority / Signature HMO',               'blocked'),
  BENESIHMNSTC: d => stat(d, 'Benefits', 'Priority / Signature HMO',               'notStarted'),
  BENESIHMB:    d => 0,

  BENEACPTTC:   d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'planned'),
  BENEACPETC:   d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'executed'),
  BENEACPPTC:   d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'passed'),
  BENEACPFTC:   d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'failed'),
  BENEACPIPTC:  d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'inProgress'),
  BENEACPBTC:   d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'blocked'),
  BENEACPNSTC:  d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'notStarted'),
  BENEACPB:     d => 0,

  BENEPRPTTC:   d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'planned'),
  BENEPRPETC:   d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'executed'),
  BENEPRPPTC:   d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'passed'),
  BENEPRPFTC:   d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'failed'),
  BENEPRPIPTC:  d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'inProgress'),
  BENEPRPBTC:   d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'blocked'),
  BENEPRPNSTC:  d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'notStarted'),
  BENEPRPB:     d => 0,

  // ── Enrollment (key: 'SIT') ────────────────────────────────────────────────
  ENROTTC:  d => d.stats.Enrollment.planned,
  ENROETC:  d => d.stats.Enrollment.executed,
  ENROEP:   d => pct(d.stats.Enrollment.executed,   d.stats.Enrollment.planned),
  ENROPTC:  d => d.stats.Enrollment.passed,
  ENROPP:   d => pct(d.stats.Enrollment.passed,     d.stats.Enrollment.executed),
  ENROFTC:  d => d.stats.Enrollment.failed,
  ENROFP:   d => pct(d.stats.Enrollment.failed,     d.stats.Enrollment.executed),
  ENROIPTC: d => d.stats.Enrollment.inProgress,
  ENROBTC:  d => d.stats.Enrollment.blocked,
  ENRONSTC: d => d.stats.Enrollment.notStarted,
  ENROB:    d => (d.bugs?.Enrollment ?? []).length,

  // ── EDI (key: 'SIT') ───────────────────────────────────────────────────────
  EDITTC:  d => d.stats.EDI.planned,
  EDIETC:  d => d.stats.EDI.executed,
  EDIEP:   d => pct(d.stats.EDI.executed,  d.stats.EDI.planned),
  EDIPTC:  d => d.stats.EDI.passed,
  EDIPP:   d => pct(d.stats.EDI.passed,    d.stats.EDI.executed),
  EDIFTC:  d => d.stats.EDI.failed,
  EDIFP:   d => pct(d.stats.EDI.failed,    d.stats.EDI.executed),
  EDIIPTC: d => d.stats.EDI.inProgress,
  EDIBTC:  d => d.stats.EDI.blocked,
  EDINSTC: d => d.stats.EDI.notStarted,
  EDIB:    d => (d.bugs?.EDI ?? []).length,

  // ── AI-refreshed narrative sentences (populated by extensions/ai-narrative.js) ──
  // Blank unless ai-narrative-input.json exists (written by a live agent run) — see that file.
  AI_OVERALL_STATUS:        d => d._aiNarratives?.overallStatus ?? '',
  AI_PDM_ITERATION_UPDATE:  d => d._aiNarratives?.pdmIterationUpdate ?? '',
  AI_PDM_DEFECT_SUMMARY:    d => d._aiNarratives?.pdmDefectSummary ?? '',
  AI_BENEFITS_UPDATE:       d => d._aiNarratives?.benefitsUpdate ?? '',
  AI_PDM_CURSORY_STATUS:    d => d._aiNarratives?.pdmCursoryStatus ?? '',
  AI_BENEFITS_STATUS:       d => d._aiNarratives?.benefitsStatus ?? '',
  AI_ENROLLMENT_STATUS:     d => d._aiNarratives?.enrollmentStatus ?? '',
  AI_EDI_STATUS:            d => d._aiNarratives?.ediStatus ?? '',
  AI_PDM_DQ_DEFECT_DETAIL_1: d => d._aiNarratives?.pdmDataQualityDetail1 ?? '',
  AI_PDM_DQ_DEFECT_DETAIL_2: d => d._aiNarratives?.pdmDataQualityDetail2 ?? '',
  AI_PDM_DM_DEFECT_DETAIL:  d => d._aiNarratives?.pdmDataMappingDetail ?? '',

};
