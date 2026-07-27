// Azure DevOps provider
//
// Exports two functions called by the orchestration in generate-report.js:
//   fetchTestStats(ws) — walks the test suite tree and returns execution counts
//   runQuery(ws, config) — runs any WIQL query defined in query-config.js QUERIES

import fetch from 'node-fetch';

const ORG     = process.env.ADO_ORG;
const PROJECT = process.env.ADO_PROJECT;
const TOKEN   = process.env.ADO_PAT;

const AUTH    = { 'Authorization': `Basic ${Buffer.from(':' + TOKEN).toString('base64')}`, 'Content-Type': 'application/json', 'Accept': 'application/json' };
const enc     = s => encodeURIComponent(s);
const baseUrl = () => `https://dev.azure.com/${ORG}/${enc(PROJECT)}/_apis`;

async function get(url) {
  const res = await fetch(url, { headers: AUTH });
  if (!res.ok) throw new Error(`ADO ${res.status}: ${url}\n${(await res.text()).slice(0, 400)}`);
  return res.json();
}

async function post(url, body) {
  const res = await fetch(url, { method: 'POST', headers: AUTH, body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`ADO ${res.status}: ${url}\n${(await res.text()).slice(0, 400)}`);
  return res.json();
}

// ─── Test stats (suite tree walk) ─────────────────────────────────────────────

async function fetchPlanSuites(planId) {
  const res = await get(`${baseUrl()}/testplan/plans/${planId}/suites?api-version=7.0&$expand=children`);
  return res.value || [];
}

function collectSuiteDescendants(rootId, allSuites) {
  const index  = Object.fromEntries(allSuites.map(s => [s.id, s]));
  const result = [];
  const queue  = [rootId];
  while (queue.length) {
    const id = queue.shift();
    const s  = index[id];
    if (!s) continue;
    result.push(id);
    for (const child of s.children || []) queue.push(child.id ?? child);
  }
  if (result.length === 1) {
    for (const s of allSuites) {
      if (s.parentSuite?.id === rootId || s.parent?.id === rootId)
        result.push(...collectSuiteDescendants(s.id, allSuites));
    }
  }
  return [...new Set(result)];
}

async function fetchSuiteTestPoints(planId, suiteId) {
  try {
    const res = await get(`${baseUrl()}/testplan/Plans/${planId}/Suites/${suiteId}/TestPoint?api-version=7.0&$top=500`);
    return res.value || [];
  } catch (e) {
    console.warn(`    ⚠  Suite ${suiteId}: ${e.message.split('\n')[0]}`);
    return [];
  }
}

function tallyOutcomes(points) {
  let planned = 0, executed = 0, passed = 0, failed = 0, notStarted = 0, inProgress = 0;
  for (const pt of points) {
    planned++;
    const o = (pt.results?.outcome || '').toLowerCase();
    if      (o === 'passed')                            { passed++;     executed++; }
    else if (o === 'failed')                            { failed++;     executed++; }
    else if (o === 'inprogress')                        { inProgress++; executed++; }
    else if (o && o !== 'none' && o !== 'unspecified')  { executed++; }
    else                                                { notStarted++; }
  }
  return { planned, executed, passed, failed, notStarted, inProgress };
}

export async function fetchTestStats({ name, planId, sitSuiteId }) {
  if (!planId || !sitSuiteId) {
    console.warn(`  ⚠  ${name}: missing planId or sitSuiteId`);
    return { planned: 0, executed: 0, passed: 0, failed: 0, notStarted: 0, inProgress: 0 };
  }
  const allSuites = await fetchPlanSuites(planId);
  const suiteIds  = collectSuiteDescendants(Number(sitSuiteId), allSuites);
  console.log(`    → ${name}: ${suiteIds.length} suite(s) under ${sitSuiteId}`);
  const totals = { planned: 0, executed: 0, passed: 0, failed: 0, notStarted: 0, inProgress: 0 };
  for (const sid of suiteIds) {
    const t = tallyOutcomes(await fetchSuiteTestPoints(planId, sid));
    for (const k of Object.keys(totals)) totals[k] += t[k];
  }
  return totals;
}

// Returns stats broken down by direct children of sitSuiteId.
// Each child's total rolls up all of its own descendants.
// Returns {} if the suite has no children (flat workstreams unaffected).
export async function fetchSubSuiteStats({ name, planId, sitSuiteId }) {
  if (!planId || !sitSuiteId) return {};
  const allSuites  = await fetchPlanSuites(planId);
  const index      = Object.fromEntries(allSuites.map(s => [s.id, s]));
  const rootId     = Number(sitSuiteId);
  const root       = index[rootId];
  if (!root) return {};

  // Direct children of the root suite
  let children = (root.children || []).map(c => index[c.id ?? c]).filter(Boolean);
  // Fallback: scan parentSuite references when $expand=children didn't populate
  if (!children.length) {
    children = allSuites.filter(s => s.parentSuite?.id === rootId || s.parent?.id === rootId);
  }
  if (!children.length) return {};

  const result = {};
  for (const child of children) {
    const descendantIds = collectSuiteDescendants(child.id, allSuites);
    const totals = { planned: 0, executed: 0, passed: 0, failed: 0, notStarted: 0, inProgress: 0 };
    for (const sid of descendantIds) {
      const t = tallyOutcomes(await fetchSuiteTestPoints(planId, sid));
      for (const k of Object.keys(totals)) totals[k] += t[k];
    }
    result[child.name] = totals;
  }
  return result;
}

// ─── Generic WIQL query runner ────────────────────────────────────────────────

// config can be one of three shapes:
//   { wiql: string, fieldMap }      — raw WIQL string (from UI / wiqlTemplate in ui-config.json)
//   { workItemType, excludeStates, includeStates?, orderBy, fields, fieldMap } — structured (config.js hardcoded)
export async function runQuery({ name, areaPath }, config) {
  if (!config) {
    console.warn(`  ⚠  ${name}: no ADO query config provided, skipping.`);
    return [];
  }

  let wiqlQuery;

  if (config.wiql) {
    // Raw WIQL — use as-is (user pasted from ADO query editor)
    wiqlQuery = config.wiql;
  } else {
    // Structured config — build WIQL from parts
    // areaPath is optional: when empty the query runs project-wide (global scope)
    const excludeClauses = (config.excludeStates || []).map(s => `[System.State] <> '${s}'`);
    const includeClauses = (config.includeStates || []).map(s => `[System.State] = '${s}'`);
    const stateClauses   = [...excludeClauses, ...includeClauses];
    const areaClause     = areaPath ? `AND [System.AreaPath] UNDER '${areaPath}'` : '';
    wiqlQuery = `SELECT ${config.fields.join(', ')}
    FROM WorkItems
    WHERE [System.WorkItemType] = '${config.workItemType}'
      AND [System.TeamProject] = '${PROJECT}'
      ${stateClauses.length ? 'AND ' + stateClauses.join(' AND ') : ''}
      ${areaClause}
    ORDER BY ${config.orderBy}`;
  }

  if (process.env.DEBUG) console.log(`\n  [DEBUG] ${name} WIQL:\n${wiqlQuery}\n`);

  let ids;
  try {
    const data = await post(`${baseUrl()}/wit/wiql?api-version=7.0`, { query: wiqlQuery });
    ids = (data.workItems || []).map(w => w.id);
  } catch (e) {
    console.warn(`  ⚠  ${name}: WIQL query failed — ${e.message.split('\n')[0]}`);
    return [];
  }

  if (!ids.length) return [];

  // ADO caps work item detail requests at 200 per call
  const items = [];
  for (let i = 0; i < ids.length; i += 200) {
    const chunk = ids.slice(i, i + 200).join(',');
    const res   = await get(`${baseUrl()}/wit/workitems?ids=${chunk}&api-version=7.0`);
    items.push(...(res.value || []));
  }

  return items.map(wi =>
    Object.fromEntries(Object.entries(config.fieldMap).map(([k, fn]) => [k, fn(wi)]))
  );
}
