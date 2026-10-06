// Buckets a workstream's currently-open defects by age (days since
// System.CreatedDate), for the "Defect Aging" table on each workstream's
// Testing Defect Report slide. Closed/Resolved defects aren't aged — once a
// defect is fixed its time-open is a burndown-chart concern, not an aging
// one, so only d.bugs.<WS> (open) feeds this, never d.closedBugs.<WS>.
//
// Buckets: 0-10 / 11-20 / 21-30 / 31+ days old — confirmed with the report
// owner (quarterly review, Oct 2026).

const DAY_MS = 24 * 60 * 60 * 1000;

function daysOpen(bug, today) {
  if (!bug.createdDate) return null;
  const created = new Date(`${bug.createdDate}T00:00:00Z`);
  return Math.floor((today - created) / DAY_MS);
}

// today defaults to now (America/New_York calendar day, matching the rest of
// the report) but is injectable for tests.
export function agingBuckets(bugs, today = new Date(new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' }) + 'T00:00:00Z')) {
  const buckets = { d0_10: 0, d11_20: 0, d21_30: 0, d31p: 0 };
  for (const bug of bugs) {
    const age = daysOpen(bug, today);
    if (age === null) continue;
    if (age <= 10) buckets.d0_10++;
    else if (age <= 20) buckets.d11_20++;
    else if (age <= 30) buckets.d21_30++;
    else buckets.d31p++;
  }
  return buckets;
}
