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

import { pdmDefectBucket, benefitsDefectPlan } from './extensions/_defect-classification.js';
import { agingBuckets } from './extensions/_defect-aging.js';
import { SIT_SCHEDULE } from './config.js';
import { computeTcsPlanning } from './extensions/_tcs-planning.js';

// Whole-number percent, except where rounding would misstate a non-complete
// (or non-zero) value as exactly 100% / 0% — those show one decimal (99.6).
const pct  = (n, d) => {
  if (!d) return 0;
  const r = Math.round((n / d) * 100);
  if (n > 0 && n < d && (r >= 100 || r <= 0)) {
    const one = Math.round((n / d) * 1000) / 10;
    return one >= 100 ? 99.9 : (one <= 0 ? 0.1 : one);
  }
  return r;
};
const sub  = (d, ws, path) => d.subStats?.[ws]?.[path] ?? {};
const stat = (d, ws, path, field) => sub(d, ws, path)[field] ?? 0;

// Sums a field across multiple sub-suite paths
const sumStat = (d, ws, paths, field) => paths.reduce((acc, p) => acc + stat(d, ws, p, field), 0);

// "Not Started" absorbs In Progress and Paused — the deck no longer shows
// those as separate buckets (report-owner direction, Oct 2026), so every
// *NSTC token folds them in here rather than reporting notStarted alone.
const nsPlus    = (d, ws, path)  => stat(d, ws, path, 'notStarted') + stat(d, ws, path, 'inProgress') + stat(d, ws, path, 'paused');
const sumNsPlus = (d, ws, paths) => sumStat(d, ws, paths, 'notStarted') + sumStat(d, ws, paths, 'inProgress') + sumStat(d, ws, paths, 'paused');

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

// As of the 09/2026 template refresh, Classic HMO NEOH and MMEGWP PPO CWRU
// MAPD also have their own permanent template rows (slide 5) alongside the
// 3 Priority plans — see BENENEOH*/BENEMMEGWP* tokens below. Excluded here
// too so _dynamic-benefits.js doesn't also try to inject a duplicate row
// for them; it now only fires for a plan beyond these 5.
export const FIXED_BENEFIT_ROW_KEYS = new Set([
  ...PRIORITY_BENEFIT_KEYS,
  'HMO / Classic HMO NEOH',
  'EGWP 1 / MMEGWP PPO CWRU MAPD',
]);

const activeBenEntries = d =>
  Object.entries(d.subStats?.Benefits ?? {})
    .filter(([key, s]) => key.split(' / ').length === 2 && s.executed > 0);

const activeBenSuites = d => activeBenEntries(d).map(([, s]) => s);

// Benefits totals now cover ALL Benefits test cases (started or not), per
// report-owner direction. Only the per-plan rows on the Benefits execution
// slide stay started-only (activeBenEntries) — _dynamic-benefits.js adds a
// "Not Started" row so that table still foots to the full total.
const activeBenSum = (d, field) => d.stats?.Benefits?.[field] ?? 0;

const activeBenNsPlus = d => activeBenSum(d, 'notStarted') + activeBenSum(d, 'inProgress') + activeBenSum(d, 'paused');

// Totals for the started-only plan rows, and the remainder (not-started plans).
export const getNotStartedBenefitsRow = d => {
  const f = k => activeBenSuites(d).reduce((a, s) => a + (s[k] ?? 0), 0);
  const planned = activeBenSum(d, 'planned') - f('planned');
  return {
    planned,
    notStarted: activeBenNsPlus(d) - (f('notStarted') + f('inProgress') + f('paused')),
    plansLeft: Object.keys(d.subStats?.Benefits ?? {}).filter(k => k.split(' / ').length === 2 && !activeBenEntries(d).some(([a]) => a === k)).length,
  };
};

// ADO sub-suite names are inconsistently cased (e.g. "MMEGWP PPO CWRU MAPD").
// Title-case each word for display in the report regardless of source casing.
const toTitleCase = s =>
  s.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());

// Active benefit plans NOT already covered by a hardcoded Priority row —
// these are what extensions/_dynamic-benefits.js inserts new rows for.
export const getExtraActiveBenefitPlans = d =>
  activeBenEntries(d)
    .filter(([key]) => !FIXED_BENEFIT_ROW_KEYS.has(key))
    .map(([key, s]) => ({ key, label: toTitleCase(key.split(' / ').pop()), stats: s }));

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
  // "Not Started" absorbs In Progress and Paused — the deck no longer shows
  // those as separate buckets (report-owner direction, Oct 2026).
  NSTC: d => activeGrandTotal(d, 'notStarted') + activeGrandTotal(d, 'inProgress') + activeGrandTotal(d, 'paused'),
  TB:   d => d.bugsBySeverity.total,

  // ── PDM — Cursory sub-suite (key: 'Cursory') ──────────────────────────────
  PDMCTTC:   d => stat(d, 'PDM', 'Cursory', 'planned'),
  PDMCETC:   d => stat(d, 'PDM', 'Cursory', 'executed'),
  PDMCEP:    d => pct(stat(d, 'PDM', 'Cursory', 'executed'),  stat(d, 'PDM', 'Cursory', 'planned')),
  PDMCPTC:   d => stat(d, 'PDM', 'Cursory', 'passed'),
  PDMCPP:    d => pct(stat(d, 'PDM', 'Cursory', 'passed'),    stat(d, 'PDM', 'Cursory', 'executed')),
  PDMCFTC:   d => stat(d, 'PDM', 'Cursory', 'failed'),
  PDMCFP:    d => pct(stat(d, 'PDM', 'Cursory', 'failed'),    stat(d, 'PDM', 'Cursory', 'executed')),
  PDMCB:     d => (d.bugs?.PDM ?? []).length,

  // ── PDM — SIT sub-suite (key: 'SIT') ──────────────────────────────────────
  PDMSITTC:   d => stat(d, 'PDM', 'SIT', 'planned'),
  PDMSITETC:  d => stat(d, 'PDM', 'SIT', 'executed'),
  PDMSITEP:   d => pct(stat(d, 'PDM', 'SIT', 'executed'),  stat(d, 'PDM', 'SIT', 'planned')),
  PDMSITPTC:  d => stat(d, 'PDM', 'SIT', 'passed'),
  PDMSITPP:   d => pct(stat(d, 'PDM', 'SIT', 'passed'),    stat(d, 'PDM', 'SIT', 'executed')),
  PDMSITFTC:  d => stat(d, 'PDM', 'SIT', 'failed'),
  PDMSITFP:   d => pct(stat(d, 'PDM', 'SIT', 'failed'),    stat(d, 'PDM', 'SIT', 'executed')),
  PDMSITBTC:  d => stat(d, 'PDM', 'SIT', 'blocked'),
  PDMSITNSTC: d => nsPlus(d, 'PDM', 'SIT'),
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
  PDMIT2BTC:   d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0', 'blocked'),
  PDMIT2NSTC:  d => nsPlus(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.0'),
  PDMIT2B:     d => 0,

  PDMIT21TTC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'planned'),
  PDMIT21ETC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'executed'),
  PDMIT21EP:   d => pct(stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'executed'), stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'planned')),
  PDMIT21PTC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'passed'),
  PDMIT21PP:   d => pct(stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'passed'),   stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'executed')),
  PDMIT21FTC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'failed'),
  PDMIT21FP:   d => pct(stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'failed'),   stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'executed')),
  PDMIT21BTC:  d => stat(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1', 'blocked'),
  PDMIT21NSTC: d => nsPlus(d, 'PDM', 'Cursory / Iteration 2 / Iteration 2.1'),
  PDMIT21B:    d => 0,

  PDMIT3TTC:   d => stat(d, 'PDM', 'Cursory / Iteration 3', 'planned'),
  PDMIT3ETC:   d => stat(d, 'PDM', 'Cursory / Iteration 3', 'executed'),
  PDMIT3EP:    d => pct(stat(d, 'PDM', 'Cursory / Iteration 3', 'executed'), stat(d, 'PDM', 'Cursory / Iteration 3', 'planned')),
  PDMIT3PTC:   d => stat(d, 'PDM', 'Cursory / Iteration 3', 'passed'),
  PDMIT3PP:    d => pct(stat(d, 'PDM', 'Cursory / Iteration 3', 'passed'),   stat(d, 'PDM', 'Cursory / Iteration 3', 'executed')),
  PDMIT3FTC:   d => stat(d, 'PDM', 'Cursory / Iteration 3', 'failed'),
  PDMIT3FP:    d => pct(stat(d, 'PDM', 'Cursory / Iteration 3', 'failed'),   stat(d, 'PDM', 'Cursory / Iteration 3', 'executed')),
  PDMIT3BTC:   d => stat(d, 'PDM', 'Cursory / Iteration 3', 'blocked'),
  PDMIT3NSTC:  d => nsPlus(d, 'PDM', 'Cursory / Iteration 3'),
  PDMIT3B:     d => 0,

  // ── PDM — Iteration 3 Data Source breakdown ────────────────────────────────
  // CPIMS has no parent suite — summed across 4 individual sub-suites
  PDMIT3CPTTC:  d => sumStat(d, 'PDM', CPIMS3_PATHS, 'planned'),
  PDMIT3CPETC:  d => sumStat(d, 'PDM', CPIMS3_PATHS, 'executed'),
  PDMIT3CPPTC:  d => sumStat(d, 'PDM', CPIMS3_PATHS, 'passed'),
  PDMIT3CPFTC:  d => sumStat(d, 'PDM', CPIMS3_PATHS, 'failed'),
  PDMIT3CPBTC:  d => sumStat(d, 'PDM', CPIMS3_PATHS, 'blocked'),
  PDMIT3CPNSTC: d => sumNsPlus(d, 'PDM', CPIMS3_PATHS),
  PDMIT3CPB:    d => 0,

  PDMIT3ANTTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'planned'),
  PDMIT3ANETC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'executed'),
  PDMIT3ANPTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'passed'),
  PDMIT3ANFTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'failed'),
  PDMIT3ANBTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Ancillary',        'blocked'),
  PDMIT3ANNSTC: d => nsPlus(d, 'PDM', 'Cursory / Iteration 3 / Ancillary'),
  PDMIT3ANB:    d => 0,

  PDMIT3ROTTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'planned'),
  PDMIT3ROETC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'executed'),
  PDMIT3ROPTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'passed'),
  PDMIT3ROFTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'failed'),
  PDMIT3ROBTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Rosters',          'blocked'),
  PDMIT3RONSTC: d => nsPlus(d, 'PDM', 'Cursory / Iteration 3 / Rosters'),
  PDMIT3ROB:    d => 0,

  PDMIT3PNTTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'planned'),
  PDMIT3PNETC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'executed'),
  PDMIT3PNPTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'passed'),
  PDMIT3PNFTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'failed'),
  PDMIT3PNBTC:  d => stat(d, 'PDM', 'Cursory / Iteration 3 / Provider Network', 'blocked'),
  PDMIT3PNNSTC: d => nsPlus(d, 'PDM', 'Cursory / Iteration 3 / Provider Network'),
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
  BENEPBBTC:   d => stat(d, 'Benefits', 'Priority', 'blocked'),
  BENEPBNSTC:  d => nsPlus(d, 'Benefits', 'Priority'),
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
  BENE26BTC:   d => activeBenSum(d, 'blocked')    - stat(d, 'Benefits', 'Priority', 'blocked'),
  BENE26NSTC:  d => activeBenNsPlus(d) - nsPlus(d, 'Benefits', 'Priority'),
  BENE26B:     d => 0,

  // ── Benefits — active total (Priority + any active non-Priority plans) ───
  // Used by slide 4's own Grand Total row (BENEPBTTC alone under-counted
  // once a non-Priority plan went active — see activeBenSum above).
  ACTIVEBENTTC:  d => activeBenSum(d, 'planned'),
  ACTIVEBENETC:  d => activeBenSum(d, 'executed'),
  ACTIVEBENEP:   d => pct(activeBenSum(d, 'executed'), activeBenSum(d, 'planned')),
  ACTIVEBENPTC:  d => activeBenSum(d, 'passed'),
  ACTIVEBENPP:   d => pct(activeBenSum(d, 'passed'),   activeBenSum(d, 'executed')),
  ACTIVEBENFTC:  d => activeBenSum(d, 'failed'),
  ACTIVEBENFP:   d => pct(activeBenSum(d, 'failed'),   activeBenSum(d, 'executed')),
  ACTIVEBENBTC:  d => activeBenSum(d, 'blocked'),
  ACTIVEBENNSTC: d => activeBenNsPlus(d),
  ACTIVEBENB:    d => (d.bugs?.Benefits ?? []).length,

  // ── Benefits — Benefit type breakdown (Priority sub-suites) ───────────────
  BENESIHMTTC:  d => stat(d, 'Benefits', 'Priority / Signature HMO',               'planned'),
  BENESIHMETC:  d => stat(d, 'Benefits', 'Priority / Signature HMO',               'executed'),
  BENESIHMPTC:  d => stat(d, 'Benefits', 'Priority / Signature HMO',               'passed'),
  BENESIHMFTC:  d => stat(d, 'Benefits', 'Priority / Signature HMO',               'failed'),
  BENESIHMBTC:  d => stat(d, 'Benefits', 'Priority / Signature HMO',               'blocked'),
  BENESIHMNSTC: d => nsPlus(d, 'Benefits', 'Priority / Signature HMO'),
  BENESIHMB:    d => 0,

  BENEACPTTC:   d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'planned'),
  BENEACPETC:   d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'executed'),
  BENEACPPTC:   d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'passed'),
  BENEACPFTC:   d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'failed'),
  BENEACPBTC:   d => stat(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)', 'blocked'),
  BENEACPNSTC:  d => nsPlus(d, 'Benefits', 'Priority / Access PPO (Premium PPO INN)'),
  BENEACPB:     d => 0,

  BENEPRPTTC:   d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'planned'),
  BENEPRPETC:   d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'executed'),
  BENEPRPPTC:   d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'passed'),
  BENEPRPFTC:   d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'failed'),
  BENEPRPBTC:   d => stat(d, 'Benefits', 'Priority / Premium PPO',                 'blocked'),
  BENEPRPNSTC:  d => nsPlus(d, 'Benefits', 'Priority / Premium PPO'),
  BENEPRPB:     d => 0,

  // ── Enrollment (key: 'SIT') ────────────────────────────────────────────────
  ENROTTC:  d => d.stats.Enrollment.planned,
  ENROETC:  d => d.stats.Enrollment.executed,
  ENROEP:   d => pct(d.stats.Enrollment.executed,   d.stats.Enrollment.planned),
  ENROPTC:  d => d.stats.Enrollment.passed,
  ENROPP:   d => pct(d.stats.Enrollment.passed,     d.stats.Enrollment.executed),
  ENROFTC:  d => d.stats.Enrollment.failed,
  ENROFP:   d => pct(d.stats.Enrollment.failed,     d.stats.Enrollment.executed),
  ENROBTC:  d => d.stats.Enrollment.blocked,
  // In Progress and Paused test cases are reported as Not Started here
  // (report-owner direction — neither has produced a result yet, so both
  // belong in the same bucket as everything else that hasn't started).
  ENRONSTC: d => d.stats.Enrollment.notStarted + d.stats.Enrollment.inProgress + d.stats.Enrollment.paused,
  ENROB:    d => (d.bugs?.Enrollment ?? []).length,

  // ── EDI (key: 'SIT') ───────────────────────────────────────────────────────
  EDITTC:  d => d.stats.EDI.planned,
  EDIETC:  d => d.stats.EDI.executed,
  EDIEP:   d => pct(d.stats.EDI.executed,  d.stats.EDI.planned),
  EDIPTC:  d => d.stats.EDI.passed,
  EDIPP:   d => pct(d.stats.EDI.passed,    d.stats.EDI.executed),
  EDIFTC:  d => d.stats.EDI.failed,
  EDIFP:   d => pct(d.stats.EDI.failed,    d.stats.EDI.executed),
  EDIBTC:  d => d.stats.EDI.blocked,
  EDINSTC: d => d.stats.EDI.notStarted + d.stats.EDI.inProgress + d.stats.EDI.paused,
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

  // ── AI-refreshed narrative — 09/2026 template additions ───────────────────
  AI_DEFECT_TRIAGE:                d => d._aiNarratives?.defectTriage ?? '',
  AI_ENROLLMENT_UPDATE:            d => d._aiNarratives?.enrollmentUpdate ?? '',
  AI_CLAIMS_UPDATE:                d => d._aiNarratives?.claimsUpdate ?? '',
  AI_PDM_BLOCKED_NOTE:             d => d._aiNarratives?.pdmBlockedNote ?? '',
  // '****' footnote marker only when there is a blocked-cases note to point at
  PDMBLOCKEDFN:                    d => d._aiNarratives?.pdmBlockedNote ? '****' : '',
  // Shown next to the PDM SIT target date while the revised date is unconfirmed
  PDMSITENDTBD:                    () => SIT_SCHEDULE.PDM?.tbd ? ' (TBD)' : '',
  AI_PDM_DEFECT_STATUS:            d => d._aiNarratives?.pdmDefectStatus ?? '',
  AI_PDM_DEFECT_INSIGHTS:          d => d._aiNarratives?.pdmDefectInsights ?? '',
  AI_PDM_BURNDOWN_INSIGHTS:        d => d._aiNarratives?.pdmBurndownInsights ?? '',
  AI_BENEFITS_DEFECT_STATUS:       d => d._aiNarratives?.benefitsDefectStatus ?? '',
  AI_BENEFITS_DEFECT_INSIGHTS:     d => d._aiNarratives?.benefitsDefectInsights ?? '',
  AI_BENEFITS_BURNDOWN_INSIGHTS:   d => d._aiNarratives?.benefitsBurndownInsights ?? '',

  // ── AI-refreshed narrative — 09/2026 Enrollment/EDI Defect Report slides ──
  AI_ENROLLMENT_DEFECT_STATUS:     d => d._aiNarratives?.enrollmentDefectStatus ?? '',
  AI_ENROLLMENT_DEFECT_INSIGHTS:   d => d._aiNarratives?.enrollmentDefectInsights ?? '',
  AI_ENROLLMENT_BURNDOWN_INSIGHTS: d => d._aiNarratives?.enrollmentBurndownInsights ?? '',
  AI_EDI_DEFECT_STATUS:            d => d._aiNarratives?.ediDefectStatus ?? '',
  AI_EDI_DEFECT_INSIGHTS:          d => d._aiNarratives?.ediDefectInsights ?? '',
  AI_EDI_BURNDOWN_INSIGHTS:        d => d._aiNarratives?.ediBurndownInsights ?? '',

  // ── Benefits — Classic HMO NEOH / MMEGWP PPO CWRU MAPD (fixed rows) ───────
  // Promoted from dynamic-row plans to permanent template rows in the
  // 09/2026 template refresh — see FIXED_BENEFIT_ROW_KEYS above.
  BENENEOHTTC:  d => stat(d, 'Benefits', 'HMO / Classic HMO NEOH', 'planned'),
  BENENEOHETC:  d => stat(d, 'Benefits', 'HMO / Classic HMO NEOH', 'executed'),
  BENENEOHPTC:  d => stat(d, 'Benefits', 'HMO / Classic HMO NEOH', 'passed'),
  BENENEOHFTC:  d => stat(d, 'Benefits', 'HMO / Classic HMO NEOH', 'failed'),
  BENENEOHBTC:  d => stat(d, 'Benefits', 'HMO / Classic HMO NEOH', 'blocked'),
  BENENEOHNSTC: d => nsPlus(d, 'Benefits', 'HMO / Classic HMO NEOH'),

  BENEMMEGWPTTC:  d => stat(d, 'Benefits', 'EGWP 1 / MMEGWP PPO CWRU MAPD', 'planned'),
  BENEMMEGWPETC:  d => stat(d, 'Benefits', 'EGWP 1 / MMEGWP PPO CWRU MAPD', 'executed'),
  BENEMMEGWPPTC:  d => stat(d, 'Benefits', 'EGWP 1 / MMEGWP PPO CWRU MAPD', 'passed'),
  BENEMMEGWPFTC:  d => stat(d, 'Benefits', 'EGWP 1 / MMEGWP PPO CWRU MAPD', 'failed'),
  BENEMMEGWPBTC:  d => stat(d, 'Benefits', 'EGWP 1 / MMEGWP PPO CWRU MAPD', 'blocked'),
  BENEMMEGWPNSTC: d => nsPlus(d, 'Benefits', 'EGWP 1 / MMEGWP PPO CWRU MAPD'),

  // ── PDM / Benefits defect breakdown by bucket (09/2026 template) ─────────
  // ADO's Area Path can't tell these apart (see _defect-classification.js
  // for why) — classified from Tags (PDM) / Title+Repro Steps (Benefits)
  // per report-owner direction. Coverage isn't 100%: a bug that doesn't
  // classify into any bucket still counts in the reliable workstream-wide
  // totals (PDMCB, ACTIVEBENB, etc.) but won't appear in any one bucket
  // below, so per-bucket sums can run slightly under the workstream total.
  ...(() => {
    const pdmBugsAll    = d => [...(d.bugs?.PDM ?? []), ...(d.closedBugs?.PDM ?? [])];
    const pdmBucketOpen   = (d, bucket) => (d.bugs?.PDM ?? []).filter(b => pdmDefectBucket(b) === bucket).length;
    const pdmBucketClosed = (d, bucket, state) => (d.closedBugs?.PDM ?? []).filter(b => pdmDefectBucket(b) === bucket && b.state === state).length;
    const pdmBucketTotal  = (d, bucket) => pdmBugsAll(d).filter(b => pdmDefectBucket(b) === bucket).length;

    const pdmBucketTokens = (prefix, bucket) => ({
      [`${prefix}DEFTOTAL`]:    d => pdmBucketTotal(d, bucket),
      [`${prefix}DEFOPEN`]:     d => pdmBucketOpen(d, bucket),
      [`${prefix}DEFCLOSED`]:   d => pdmBucketClosed(d, bucket, 'Closed'),
      [`${prefix}DEFRESOLVED`]: d => pdmBucketClosed(d, bucket, 'Resolved'),
    });

    const benefitsBugsAll    = d => [...(d.bugs?.Benefits ?? []), ...(d.closedBugs?.Benefits ?? [])];
    const benefitsPlanOpen   = (d, plan) => (d.bugs?.Benefits ?? []).filter(b => benefitsDefectPlan(b) === plan).length;
    const benefitsPlanClosed = (d, plan, state) => (d.closedBugs?.Benefits ?? []).filter(b => benefitsDefectPlan(b) === plan && b.state === state).length;
    const benefitsPlanTotal  = (d, plan) => benefitsBugsAll(d).filter(b => benefitsDefectPlan(b) === plan).length;

    const benefitsPlanTokens = (prefix, plan) => ({
      [`${prefix}DEFTOTAL`]:    d => benefitsPlanTotal(d, plan),
      [`${prefix}DEFOPEN`]:     d => benefitsPlanOpen(d, plan),
      [`${prefix}DEFCLOSED`]:   d => benefitsPlanClosed(d, plan, 'Closed'),
      [`${prefix}DEFRESOLVED`]: d => benefitsPlanClosed(d, plan, 'Resolved'),
    });

    // ── "Testing Defect Summary by Severity" tables (09/2026 template
    // addition, all 4 workstreams) ──────────────────────────────────────────
    // Same open/closed/resolved/total shape as the bucket/plan tokens above,
    // but split by ADO Severity instead — no classification heuristic
    // needed, severity is a native ADO field.
    const SEVERITY_LEVELS = [
      ['CRIT', '1 - Critical'],
      ['HIGH', '2 - High'],
      ['MED',  '3 - Medium'],
      ['LOW',  '4 - Low'],
    ];
    const severityTokens = (prefix, ws) => {
      const open   = d => d.bugs?.[ws] ?? [];
      const closed = d => d.closedBugs?.[ws] ?? [];
      const all    = d => [...open(d), ...closed(d)];
      const tokens = {};
      for (const [key, sev] of SEVERITY_LEVELS) {
        tokens[`${prefix}SEV${key}TOTAL`]    = d => all(d).filter(b => b.severity === sev).length;
        tokens[`${prefix}SEV${key}OPEN`]     = d => open(d).filter(b => b.severity === sev).length;
        tokens[`${prefix}SEV${key}CLOSED`]   = d => closed(d).filter(b => b.severity === sev && b.state === 'Closed').length;
        tokens[`${prefix}SEV${key}RESOLVED`] = d => closed(d).filter(b => b.severity === sev && b.state === 'Resolved').length;
      }
      return tokens;
    };

    return {
      ...pdmBucketTokens('PDMIT2',  'Iteration 2'),
      ...pdmBucketTokens('PDMIT21', 'Iteration 2.1'),
      ...pdmBucketTokens('PDMIT3',  'Iteration 3'),
      ...pdmBucketTokens('PDMSIT',  'SIT'),

      // Cursory-phase (any numbered iteration) open defects, for the
      // executive-summary breakdown table's "PDM - Cursory" row.
      PDMCURSORYDEFOPEN: d => pdmBucketOpen(d, 'Iteration 2') + pdmBucketOpen(d, 'Iteration 2.1') + pdmBucketOpen(d, 'Iteration 3'),

      PDMDEFTOTAL:         d => pdmBugsAll(d).length,
      PDMDEFCLOSEDTOTAL:   d => (d.closedBugs?.PDM ?? []).filter(b => b.state === 'Closed').length,
      PDMDEFRESOLVEDTOTAL: d => (d.closedBugs?.PDM ?? []).filter(b => b.state === 'Resolved').length,

      ...benefitsPlanTokens('BENESIHM',    'Signature HMO'),
      ...benefitsPlanTokens('BENEACP',     'Access PPO'),
      ...benefitsPlanTokens('BENEPRP',     'Premium PPO'),
      ...benefitsPlanTokens('BENENEOH',    'Classic HMO NEOH'),
      ...benefitsPlanTokens('BENEMMEGWP',  'MMEGWP PPO CWRU MAPD'),

      // Combined open-defect counts for the exec-summary/slide-5 aggregate
      // rows ("SIT – Priority Benefit" = the 3 priority plans; "SIT – 2026
      // Benefits" = the other 2 active plans).
      BENEPBDEF:  d => benefitsPlanOpen(d, 'Signature HMO') + benefitsPlanOpen(d, 'Access PPO') + benefitsPlanOpen(d, 'Premium PPO'),
      BENE26DEF:  d => benefitsPlanOpen(d, 'Classic HMO NEOH') + benefitsPlanOpen(d, 'MMEGWP PPO CWRU MAPD'),

      BENEDEFTOTAL:         d => benefitsBugsAll(d).length,
      BENEDEFCLOSEDTOTAL:   d => (d.closedBugs?.Benefits ?? []).filter(b => b.state === 'Closed').length,
      BENEDEFRESOLVEDTOTAL: d => (d.closedBugs?.Benefits ?? []).filter(b => b.state === 'Resolved').length,

      // Enrollment/EDI have no bucket/plan breakdown — the whole workstream
      // is a single "SIT" row on their (new, 09/2026) Testing Defect Report
      // slide, so the same 3 tokens feed both the "SIT" row and the "TOTAL"
      // row directly beneath it (intentionally identical values).
      ENRODEFTOTAL:         d => (d.bugs?.Enrollment ?? []).length + (d.closedBugs?.Enrollment ?? []).length,
      ENRODEFCLOSEDTOTAL:   d => (d.closedBugs?.Enrollment ?? []).filter(b => b.state === 'Closed').length,
      ENRODEFRESOLVEDTOTAL: d => (d.closedBugs?.Enrollment ?? []).filter(b => b.state === 'Resolved').length,

      EDIDEFTOTAL:         d => (d.bugs?.EDI ?? []).length + (d.closedBugs?.EDI ?? []).length,
      EDIDEFCLOSEDTOTAL:   d => (d.closedBugs?.EDI ?? []).filter(b => b.state === 'Closed').length,
      EDIDEFRESOLVEDTOTAL: d => (d.closedBugs?.EDI ?? []).filter(b => b.state === 'Resolved').length,

      ...severityTokens('PDM',  'PDM'),
      ...severityTokens('BENE', 'Benefits'),
      ...severityTokens('ENRO', 'Enrollment'),
      ...severityTokens('EDI',  'EDI'),
    };
  })(),

  // ── Defect Aging (10/2026 template addition, all 4 workstreams) ──────────
  // Open-defect counts bucketed by days-since-logged — see
  // extensions/_defect-aging.js. Only d.bugs.<WS> (currently open) feeds
  // this; closed/resolved defects aren't "aged".
  ...(() => {
    const agingTokens = (prefix, ws) => ({
      [`${prefix}AGE010`]:  d => agingBuckets(d.bugs?.[ws] ?? []).d0_10,
      [`${prefix}AGE1120`]: d => agingBuckets(d.bugs?.[ws] ?? []).d11_20,
      [`${prefix}AGE2130`]: d => agingBuckets(d.bugs?.[ws] ?? []).d21_30,
      [`${prefix}AGE31P`]:  d => agingBuckets(d.bugs?.[ws] ?? []).d31p,
    });
    return {
      ...agingTokens('PDM',  'PDM'),
      ...agingTokens('BENE', 'Benefits'),
      ...agingTokens('ENRO', 'Enrollment'),
      ...agingTokens('EDI',  'EDI'),
    };
  })(),

  // ── Test Case Planning Summary slide (10/2026 template addition) ─────────
  // One row per workstream: ADO test case count, a pacing-derived "planned
  // execution to date" target (+ % of total), the committed SIT end date,
  // actual execution to date (+ % of total), the delta between actual and
  // planned, a Status cell (colored by extensions/_tcs-status.js — TCSSTATUS
  // always resolves to '' here, it's purely a cell-locator token for that
  // extension), and a blank Comments cell left for manual fill-in each week.
  // Planned/actual/delta math lives in extensions/_tcs-planning.js (shared
  // with _tcs-status.js so the color and the numbers never disagree).
  ...(() => {
    const fmtDate = iso => new Date(`${iso}T00:00:00Z`).toLocaleDateString('en-US', { timeZone: 'UTC', month: '2-digit', day: '2-digit', year: 'numeric' });

    // ADO totals per workstream — Benefits uses the same active-plan filter
    // (activeBenSum) as the rest of the report rather than raw stats.Benefits,
    // so this stays consistent with every other Benefits figure in the deck.
    // PDM uses the SIT suite only (Cursory is wound down and excluded).
    // Benefits counts ALL plans (started or not), per report-owner direction.
    const adoTotal    = (d, ws) => ws === 'PDM' ? (d.subStats?.PDM?.SIT?.planned ?? 0)  : d.stats[ws].planned;
    const adoExecuted = (d, ws) => ws === 'PDM' ? (d.subStats?.PDM?.SIT?.executed ?? 0) : d.stats[ws].executed;
    const planning    = (d, ws) => computeTcsPlanning(adoTotal(d, ws), adoExecuted(d, ws), ws);

    const tcsTokens = (prefix, ws) => ({
      [`${prefix}TCSADO`]:         d => adoTotal(d, ws),
      [`${prefix}TCSPLANNED`]:     d => planning(d, ws).planned,
      [`${prefix}TCSPLANNEDPCT`]:  d => planning(d, ws).plannedPct,
      [`${prefix}TCSSITEND`]:      () => SIT_SCHEDULE[ws] ? fmtDate(SIT_SCHEDULE[ws].end) : '',
      [`${prefix}TCSACTUAL`]:      d => adoExecuted(d, ws),
      [`${prefix}TCSACTUALPCT`]:   d => planning(d, ws).actualPct,
      [`${prefix}TCSDELTA`]:       d => { const v = planning(d, ws).delta; return v > 0 ? `+${v}` : `${v}`; },
      [`${prefix}TCSSTATUS`]:      () => '',
      [`${prefix}TCSCOMMENTS`]:    () => '',
    });

    // Overall expected execution % across the four SIT workstreams (sum of
    // pacing-derived planned-to-date / sum of ADO SIT cases) — the "expected
    // execution rate" quoted in the exec summary and under the planning table.
    const WSS = ['PDM', 'Benefits', 'Enrollment', 'EDI'];
    const expectedPct = d => {
      const tot = WSS.reduce((a, ws) => a + adoTotal(d, ws), 0);
      const pl  = WSS.reduce((a, ws) => a + planning(d, ws).planned, 0);
      return tot ? Math.round((pl / tot) * 100) : 0;
    };
    const asOfShort = () => new Date().toLocaleDateString('en-US', { timeZone: 'America/New_York', month: 'numeric', day: 'numeric' });

    return {
      TCSEXPECTEDPCT: d => expectedPct(d),
      TCSOVERALLNOTE: d => {
        const behind = WSS.filter(ws => planning(d, ws).delta < 0).length;
        const lead = behind === WSS.length ? 'Overall testing is behind for all workstreams.'
                   : behind === 0 ? 'Overall testing is on or ahead of plan for all workstreams.'
                   : `Overall testing is behind for ${behind} of ${WSS.length} workstreams.`;
        return `${lead} SIT cases are expected to be ${expectedPct(d)}% executed as of ${asOfShort()}.`;
      },
      ...tcsTokens('PDM',  'PDM'),
      ...tcsTokens('BENE', 'Benefits'),
      ...tcsTokens('ENRO', 'Enrollment'),
      ...tcsTokens('EDI',  'EDI'),
    };
  })(),

};
