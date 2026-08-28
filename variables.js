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

// Each top-level workstream's executive-summary total only counts the
// sub-suites broken out on that workstream's detail slide — NOT every ADO
// sub-suite under it. PDM happens to show all 4 of its sub-suites; Enrollment
// and EDI each have sub-suites in ADO (Member/Dependent Enrollment, 834/837
// Transactions) that aren't reported yet and are intentionally excluded here.
const PDM_PATHS  = ['Plan Configuration', 'Benefit Rules', 'Rate Management', 'Network Management'];
const ENRL_PATHS = ['Open Enrollment', 'COBRA Enrollment'];
// 'Trading Partner Setup' is nested under an intermediate 'EDI SUBSSSS' suite
// in ADO rather than directly under EDI — this is its full breadcrumb path.
const EDI_PATHS  = ['File Processing', 'EDI SUBSSSS / Trading Partner Setup'];

const wsSum = (d, ws, field) => {
  const paths = ws === 'PDM' ? PDM_PATHS : ws === 'Enrollment' ? ENRL_PATHS : EDI_PATHS;
  return sumStat(d, ws, paths, field);
};

// Report-wide grand total = sum of the three filtered workstream totals above.
const reportSum = (d, field) =>
  wsSum(d, 'PDM', field) + wsSum(d, 'Enrollment', field) + wsSum(d, 'EDI', field);

export const VARIABLE_MAP = {

  // ── Report Date ────────────────────────────────────────────────────────────
  Date: () => new Date().toLocaleDateString('en-US', {
    timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: 'numeric',
  }),

  // ── Grand Totals (Executive Summary + stat boxes, slide 2) ────────────────
  TTC:  d => reportSum(d, 'planned'),
  ETC:  d => reportSum(d, 'executed'),
  EP:   d => pct(reportSum(d, 'executed'), reportSum(d, 'planned')),
  PTC:  d => reportSum(d, 'passed'),
  PP:   d => pct(reportSum(d, 'passed'),   reportSum(d, 'executed')),
  FTC:  d => reportSum(d, 'failed'),
  FP:   d => pct(reportSum(d, 'failed'),   reportSum(d, 'executed')),
  IPTC: d => reportSum(d, 'inProgress'),
  NSTC: d => reportSum(d, 'notStarted'),
  TB:   d => d.bugsBySeverity.total,

  // ── PDM — workstream total (slides 2 & 3) ─────────────────────────────────
  PDMTTC:  d => wsSum(d, 'PDM', 'planned'),
  PDMETC:  d => wsSum(d, 'PDM', 'executed'),
  PDMEP:   d => pct(wsSum(d, 'PDM', 'executed'), wsSum(d, 'PDM', 'planned')),
  PDMPTC:  d => wsSum(d, 'PDM', 'passed'),
  PDMPP:   d => pct(wsSum(d, 'PDM', 'passed'),   wsSum(d, 'PDM', 'executed')),
  PDMFTC:  d => wsSum(d, 'PDM', 'failed'),
  PDMFP:   d => pct(wsSum(d, 'PDM', 'failed'),   wsSum(d, 'PDM', 'executed')),
  PDMIPTC: d => wsSum(d, 'PDM', 'inProgress'),
  PDMBTC:  d => wsSum(d, 'PDM', 'blocked'),
  PDMNSTC: d => wsSum(d, 'PDM', 'notStarted'),
  PDMB:    d => (d.bugs?.PDM ?? []).length,

  // ── PDM — sub-suites (slide 3 detail table) ───────────────────────────────
  // Per-sub-suite defect counts are hardcoded to 0: ADO only has an area path
  // at the PDM level, not per sub-suite, so bugs can't be split out further.
  NETMTTC:  d => stat(d, 'PDM', 'Network Management', 'planned'),
  NETMETC:  d => stat(d, 'PDM', 'Network Management', 'executed'),
  NETMPTC:  d => stat(d, 'PDM', 'Network Management', 'passed'),
  NETMFTC:  d => stat(d, 'PDM', 'Network Management', 'failed'),
  NETMIPTC: d => stat(d, 'PDM', 'Network Management', 'inProgress'),
  NETMBTC:  d => stat(d, 'PDM', 'Network Management', 'blocked'),
  NETMNSTC: d => stat(d, 'PDM', 'Network Management', 'notStarted'),
  NETMB:    d => 0,

  RATMTTC:  d => stat(d, 'PDM', 'Rate Management', 'planned'),
  RATMETC:  d => stat(d, 'PDM', 'Rate Management', 'executed'),
  RATMPTC:  d => stat(d, 'PDM', 'Rate Management', 'passed'),
  RATMFTC:  d => stat(d, 'PDM', 'Rate Management', 'failed'),
  RATMIPTC: d => stat(d, 'PDM', 'Rate Management', 'inProgress'),
  RATMBTC:  d => stat(d, 'PDM', 'Rate Management', 'blocked'),
  RATMNSTC: d => stat(d, 'PDM', 'Rate Management', 'notStarted'),
  RATMB:    d => 0,

  BNFRTTC:  d => stat(d, 'PDM', 'Benefit Rules', 'planned'),
  BNFRETC:  d => stat(d, 'PDM', 'Benefit Rules', 'executed'),
  BNFRPTC:  d => stat(d, 'PDM', 'Benefit Rules', 'passed'),
  BNFRFTC:  d => stat(d, 'PDM', 'Benefit Rules', 'failed'),
  BNFRIPTC: d => stat(d, 'PDM', 'Benefit Rules', 'inProgress'),
  BNFRBTC:  d => stat(d, 'PDM', 'Benefit Rules', 'blocked'),
  BNFRNSTC: d => stat(d, 'PDM', 'Benefit Rules', 'notStarted'),
  BNFRB:    d => 0,

  PLNCTTC:  d => stat(d, 'PDM', 'Plan Configuration', 'planned'),
  PLNCETC:  d => stat(d, 'PDM', 'Plan Configuration', 'executed'),
  PLNCPTC:  d => stat(d, 'PDM', 'Plan Configuration', 'passed'),
  PLNCFTC:  d => stat(d, 'PDM', 'Plan Configuration', 'failed'),
  PLNCIPTC: d => stat(d, 'PDM', 'Plan Configuration', 'inProgress'),
  PLNCBTC:  d => stat(d, 'PDM', 'Plan Configuration', 'blocked'),
  PLNCNSTC: d => stat(d, 'PDM', 'Plan Configuration', 'notStarted'),
  PLNCB:    d => 0,

  // ── Enrollment — workstream total (slides 2 & 5) ──────────────────────────
  ENRLTTC:  d => wsSum(d, 'Enrollment', 'planned'),
  ENRLETC:  d => wsSum(d, 'Enrollment', 'executed'),
  ENRLEP:   d => pct(wsSum(d, 'Enrollment', 'executed'), wsSum(d, 'Enrollment', 'planned')),
  ENRLPTC:  d => wsSum(d, 'Enrollment', 'passed'),
  ENRLPP:   d => pct(wsSum(d, 'Enrollment', 'passed'),   wsSum(d, 'Enrollment', 'executed')),
  ENRLFTC:  d => wsSum(d, 'Enrollment', 'failed'),
  ENRLFP:   d => pct(wsSum(d, 'Enrollment', 'failed'),   wsSum(d, 'Enrollment', 'executed')),
  ENRLIPTC: d => wsSum(d, 'Enrollment', 'inProgress'),
  ENRLBTC:  d => wsSum(d, 'Enrollment', 'blocked'),
  ENRLNSTC: d => wsSum(d, 'Enrollment', 'notStarted'),
  ENRLB:    d => (d.bugs?.Enrollment ?? []).length,

  // ── Enrollment — sub-suites (slide 5 detail table) ────────────────────────
  COBRTTC:  d => stat(d, 'Enrollment', 'COBRA Enrollment', 'planned'),
  COBRETC:  d => stat(d, 'Enrollment', 'COBRA Enrollment', 'executed'),
  COBRPTC:  d => stat(d, 'Enrollment', 'COBRA Enrollment', 'passed'),
  COBRFTC:  d => stat(d, 'Enrollment', 'COBRA Enrollment', 'failed'),
  COBRIPTC: d => stat(d, 'Enrollment', 'COBRA Enrollment', 'inProgress'),
  COBRBTC:  d => stat(d, 'Enrollment', 'COBRA Enrollment', 'blocked'),
  COBRNSTC: d => stat(d, 'Enrollment', 'COBRA Enrollment', 'notStarted'),
  COBRB:    d => 0,

  OPENTTC:  d => stat(d, 'Enrollment', 'Open Enrollment', 'planned'),
  OPENETC:  d => stat(d, 'Enrollment', 'Open Enrollment', 'executed'),
  OPENPTC:  d => stat(d, 'Enrollment', 'Open Enrollment', 'passed'),
  OPENFTC:  d => stat(d, 'Enrollment', 'Open Enrollment', 'failed'),
  OPENIPTC: d => stat(d, 'Enrollment', 'Open Enrollment', 'inProgress'),
  OPENBTC:  d => stat(d, 'Enrollment', 'Open Enrollment', 'blocked'),
  OPENNSTC: d => stat(d, 'Enrollment', 'Open Enrollment', 'notStarted'),
  OPENB:    d => 0,

  // ── EDI — workstream total (slides 2 & 4) ─────────────────────────────────
  EDITTC:  d => wsSum(d, 'EDI', 'planned'),
  EDIETC:  d => wsSum(d, 'EDI', 'executed'),
  EDIEP:   d => pct(wsSum(d, 'EDI', 'executed'), wsSum(d, 'EDI', 'planned')),
  EDIPTC:  d => wsSum(d, 'EDI', 'passed'),
  EDIPP:   d => pct(wsSum(d, 'EDI', 'passed'),   wsSum(d, 'EDI', 'executed')),
  EDIFTC:  d => wsSum(d, 'EDI', 'failed'),
  EDIFP:   d => pct(wsSum(d, 'EDI', 'failed'),   wsSum(d, 'EDI', 'executed')),
  EDIIPTC: d => wsSum(d, 'EDI', 'inProgress'),
  EDIBTC:  d => wsSum(d, 'EDI', 'blocked'),
  EDINSTC: d => wsSum(d, 'EDI', 'notStarted'),
  EDIB:    d => (d.bugs?.EDI ?? []).length,

  // ── EDI — sub-suites (slide 4 detail table) ───────────────────────────────
  FILPTTC:  d => stat(d, 'EDI', 'File Processing', 'planned'),
  FILPETC:  d => stat(d, 'EDI', 'File Processing', 'executed'),
  FILPPTC:  d => stat(d, 'EDI', 'File Processing', 'passed'),
  FILPFTC:  d => stat(d, 'EDI', 'File Processing', 'failed'),
  FILPIPTC: d => stat(d, 'EDI', 'File Processing', 'inProgress'),
  FILPBTC:  d => stat(d, 'EDI', 'File Processing', 'blocked'),
  FILPNSTC: d => stat(d, 'EDI', 'File Processing', 'notStarted'),
  FILPB:    d => 0,

  // Nested two levels deep in ADO under an intermediate 'EDI SUBSSSS' suite.
  TPSTTTC:  d => stat(d, 'EDI', 'EDI SUBSSSS / Trading Partner Setup', 'planned'),
  TPSTETC:  d => stat(d, 'EDI', 'EDI SUBSSSS / Trading Partner Setup', 'executed'),
  TPSTPTC:  d => stat(d, 'EDI', 'EDI SUBSSSS / Trading Partner Setup', 'passed'),
  TPSTFTC:  d => stat(d, 'EDI', 'EDI SUBSSSS / Trading Partner Setup', 'failed'),
  TPSTIPTC: d => stat(d, 'EDI', 'EDI SUBSSSS / Trading Partner Setup', 'inProgress'),
  TPSTBTC:  d => stat(d, 'EDI', 'EDI SUBSSSS / Trading Partner Setup', 'blocked'),
  TPSTNSTC: d => stat(d, 'EDI', 'EDI SUBSSSS / Trading Partner Setup', 'notStarted'),
  TPSTB:    d => 0,

  // ── AI-refreshed narrative sentences (populated by extensions/ai-narrative.js) ──
  // Blank unless ai-narrative-input.json exists (written by a live agent run) — see that file.
  AI_OVERALL_STATUS:       d => d._aiNarratives?.overallStatus ?? '',
  AI_PDM_DEFECT_SUMMARY:   d => d._aiNarratives?.pdmDefectSummary ?? '',
  AI_DEFECT_TRIAGE_STATUS: d => d._aiNarratives?.defectTriageStatus ?? '',
  AI_PDM_STATUS:           d => d._aiNarratives?.pdmStatus ?? '',
  AI_EDI_STATUS:           d => d._aiNarratives?.ediStatus ?? '',
  AI_ENROLLMENT_STATUS:    d => d._aiNarratives?.enrollmentStatus ?? '',

};
