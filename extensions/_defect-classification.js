// MMO-specific: attributes a PDM or Benefits bug to the sub-bucket the new
// template's tables break defects down by (PDM: Cursory Iteration 2 / 2.1 /
// 3 / SIT; Benefits: Signature HMO / Access PPO / Premium PPO / Classic HMO
// NEOH / MMEGWP PPO CWRU MAPD). ADO's Area Path isn't granular enough for
// this — every PDM bug shares one Area Path, and same for Benefits — so
// this uses System.Tags (PDM) and Title/Repro Steps keyword matching
// (Benefits) per direction from the report owner. Isolated here rather than
// folded into variables.js since it's MMO-specific triage logic, not a
// generic reporting concern.
//
// KNOWN LIMITATION: neither classifier is 100% — a handful of PDM bugs carry
// only a generic "Cursory" tag with no iteration number (or no recognized
// tag at all), and a small number of Benefits bugs don't mention any plan
// name in their title or repro steps. Those bugs are excluded from the
// per-bucket breakdown (returns null) but are still counted correctly in
// the workstream-wide totals elsewhere in variables.js.

export function pdmDefectBucket(bug) {
  const tags = (bug.tags || '').split(';').map(t => t.trim());
  if (tags.includes('SIT') || tags.includes('Missing Functionality') || tags.includes('Enhancement')) return 'SIT';
  if (tags.includes('Iteration 2.1')) return 'Iteration 2.1';
  if (tags.includes('Iteration 2'))   return 'Iteration 2';
  if (tags.includes('Iteration 3'))   return 'Iteration 3';
  return null;
}

// Priority order matters — checked most-specific first so a repro step that
// happens to mention multiple plans (common when steps reference a similar
// test case on another plan for comparison) resolves deterministically
// instead of being left ambiguous.
const BENEFITS_PLAN_PATTERNS = [
  ['MMEGWP PPO CWRU MAPD', /cwru/i],
  ['Classic HMO NEOH',     /classic\s*hmo|neoh/i],
  ['Signature HMO',        /signature/i],
  ['Access PPO',           /access\s*ppo|premium\s*ppo\s*inn|ppo\s*inn/i],
  ['Premium PPO',          /premium\s*ppo/i],
];

export function benefitsDefectPlan(bug) {
  const title = bug.title || '';
  for (const [label, pattern] of BENEFITS_PLAN_PATTERNS) {
    if (pattern.test(title)) return label;
  }
  const combined = `${title} ${bug.reproSteps || ''}`;
  for (const [label, pattern] of BENEFITS_PLAN_PATTERNS) {
    if (pattern.test(combined)) return label;
  }
  return null;
}
