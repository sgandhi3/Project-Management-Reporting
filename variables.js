// This file controls what gets injected into each {{TOKEN}} in your PowerPoint template.
// Each entry is: TOKEN_NAME: d => <expression that reads from the data object>
//
// The data object `d` has this shape:
//   d.stats.<WorkstreamName>        — { planned, executed, passed, failed, notStarted, inProgress }
//   d.<queryKey>.<WorkstreamName>   — array of items from that query (e.g. d.bugs.PDM, d.closedBugs.PDM)
//   d.consolidatedData              — stats summed across all workstreams
//   d.bugsBySeverityPriority        — { total, severity: { total, sev1–4 }, priority: { pri1–4 } }
//
// WorkstreamName must match the `name` field in the WORKSTREAMS array in config.js.

export const VARIABLE_MAP = {

  // PDM
  PDMTTC:    d => d.stats.PDM.planned,
  PDMETC:    d => d.stats.PDM.executed,
  PDMNSTC:   d => d.stats.PDM.notStarted,
  PDMIPTC:   d => d.stats.PDM.inProgress,
  PDMPTC:    d => d.stats.PDM.passed,
  PDMFTC:    d => d.stats.PDM.failed,
  PDMTB:     d => d.bugs.PDM.length,

  // Benefits
  BTTC:      d => d.stats.Benefits.planned,
  BETC:      d => d.stats.Benefits.executed,
  BNSTC:     d => d.stats.Benefits.notStarted,
  BIPTC:     d => d.stats.Benefits.inProgress,
  BPTC:      d => d.stats.Benefits.passed,
  BFTC:      d => d.stats.Benefits.failed,
  BTB:       d => d.bugs.Benefits.length,

  // Enrollment
  ETTC:      d => d.stats.Enrollment.planned,
  EETC:      d => d.stats.Enrollment.executed,
  ENSTC:     d => d.stats.Enrollment.notStarted,
  EIPTC:     d => d.stats.Enrollment.inProgress,
  EPTC:      d => d.stats.Enrollment.passed,
  EFTC:      d => d.stats.Enrollment.failed,
  ETB:       d => d.bugs.Enrollment.length,

  // EDI
  EDITTC:    d => d.stats.EDI.planned,
  EDIETC:    d => d.stats.EDI.executed,
  EDINSTC:   d => d.stats.EDI.notStarted,
  EDIIPTC:   d => d.stats.EDI.inProgress,
  EDIPTC:    d => d.stats.EDI.passed,
  EDIFTC:    d => d.stats.EDI.failed,
  EDITB:     d => d.bugs.EDI.length,

  // Overall totals across all workstreams
  TTC:       d => d.consolidatedData.planned,
  ETC:       d => d.consolidatedData.executed,
  NSTC:      d => d.consolidatedData.notStarted,
  IPTC:      d => d.consolidatedData.inProgress,
  PTC:       d => d.consolidatedData.passed,
  FTC:       d => d.consolidatedData.failed,
  TB:        d => d.bugsBySeverityPriority.total,

  // Calculated percentages — guard against divide-by-zero if nothing has run yet
  PP:        d => d.consolidatedData.executed ? Math.round((d.consolidatedData.passed / d.consolidatedData.executed) * 100) : 0,
  FP:        d => d.consolidatedData.executed ? Math.round((d.consolidatedData.failed / d.consolidatedData.executed) * 100) : 0,

  // Report date in EST, formatted MM/DD/YYYY
  Date:      () => new Date().toLocaleDateString('en-US', { timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: 'numeric' }),

};
