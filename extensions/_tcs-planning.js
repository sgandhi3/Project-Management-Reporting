// Shared by variables.js (text tokens) and _tcs-status.js (the Status cell's
// fill color) for the Test Case Planning Summary slide — kept in one place
// so the two never drift apart on what "planned to date" or "on track"
// means.
import { SIT_SCHEDULE } from '../config.js';

const DAY_MS = 24 * 60 * 60 * 1000;
const today = () => new Date(new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' }) + 'T00:00:00Z');

// "Expected executed by now" = (total cases / total SIT-window days) *
// elapsed days since the window's start, capped to the window.
function plannedToDate(totalCases, schedule) {
  if (!schedule) return 0;
  const start = new Date(`${schedule.start}T00:00:00Z`);
  const end   = new Date(`${schedule.end}T00:00:00Z`);
  const totalDays = (end - start) / DAY_MS;
  if (totalDays <= 0) return totalCases;
  const elapsedDays = Math.min(Math.max((today() - start) / DAY_MS, 0), totalDays);
  return Math.round((totalCases / totalDays) * elapsedDays);
}

// Green: at/ahead of pace (actual >= 90% of planned-to-date). Yellow:
// somewhat behind (50-90%). Red: significantly behind (<50%). A SIT window
// that hasn't started yet (planned-to-date == 0) can't be "behind" by
// definition, so it's green regardless of actual.
export function tcsStatusColor(planned, actual) {
  if (planned <= 0) return '2E7D32';
  const ratio = actual / planned;
  if (ratio >= 0.9) return '2E7D32';
  if (ratio >= 0.5) return 'F9A825';
  return 'C62828';
}

export function computeTcsPlanning(ado, actual, ws) {
  const planned = plannedToDate(ado, SIT_SCHEDULE[ws]);
  const plannedPct = ado ? Math.round((planned / ado) * 100) : 0;
  const actualPct  = ado ? Math.round((actual  / ado) * 100) : 0;
  const delta = actual - planned;
  return { planned, plannedPct, actualPct, delta, color: SIT_SCHEDULE[ws]?.statusColor ?? tcsStatusColor(planned, actual) };
}
