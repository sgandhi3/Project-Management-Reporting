// Jira + Zephyr Scale provider
//
// Zephyr Scale is a Jira add-on — test execution comes from Zephyr cycle executions,
// everything else (bugs, tasks, etc.) comes from Jira via JQL.
//
// Required env vars: JIRA_DOMAIN, JIRA_EMAIL, JIRA_API_TOKEN, ZEPHYR_TOKEN
// Per-workstream config in config.js: projectKey (Jira), testCycleKey (Zephyr cycle key)
//
// Exports:
//   fetchTestStats(ws) — fetches cycle executions from Zephyr Scale
//   runQuery(ws, jql)  — runs any JQL string defined in config.js QUERIES

import fetch from 'node-fetch';
import { ZEPHYR_STATUS_MAP, JIRA_PRIORITY_MAP } from '../config.js';

const ZEPHYR_TOKEN = process.env.ZEPHYR_TOKEN;
const JIRA_DOMAIN  = process.env.JIRA_DOMAIN;
const JIRA_EMAIL   = process.env.JIRA_EMAIL;
const JIRA_TOKEN   = process.env.JIRA_API_TOKEN;

const JIRA_AUTH = {
  'Authorization': `Basic ${Buffer.from(`${JIRA_EMAIL}:${JIRA_TOKEN}`).toString('base64')}`,
  'Content-Type':  'application/json',
  'Accept':        'application/json',
};

// ─── Zephyr API ───────────────────────────────────────────────────────────────

async function zephyrGet(path) {
  const res = await fetch(`https://api.zephyrscale.smartbear.com/v2${path}`, {
    headers: { 'Authorization': `Bearer ${ZEPHYR_TOKEN}`, 'Accept': 'application/json' },
  });
  if (!res.ok) throw new Error(`Zephyr ${res.status}: ${path}\n${(await res.text()).slice(0, 400)}`);
  return res.json();
}

async function fetchCycleExecutions(cycleKey) {
  const executions = [];
  let startAt = 0;
  while (true) {
    const data = await zephyrGet(`/testexecutions?testCycle=${cycleKey}&maxResults=100&startAt=${startAt}`);
    executions.push(...(data.values || []));
    if (executions.length >= (data.total || 0) || !(data.values || []).length) break;
    startAt += data.values.length;
  }
  return executions;
}

// ─── Jira API ─────────────────────────────────────────────────────────────────

async function jiraSearch(jql, fields) {
  const allIssues = [];
  let startAt = 0;
  while (true) {
    const res = await fetch(`https://${JIRA_DOMAIN}.atlassian.net/rest/api/3/search`, {
      method: 'POST', headers: JIRA_AUTH,
      body:   JSON.stringify({ jql, fields, maxResults: 100, startAt }),
    });
    if (!res.ok) throw new Error(`Jira ${res.status}: ${(await res.text()).slice(0, 400)}`);
    const data = await res.json();
    allIssues.push(...(data.issues || []));
    if (allIssues.length >= data.total || !(data.issues || []).length) break;
    startAt += data.issues.length;
  }
  return allIssues;
}

function shapeIssue(issue) {
  const priority = (issue.fields.priority?.name || '').toLowerCase();
  return {
    id:          issue.key,
    title:       issue.fields.summary || '',
    severity:    JIRA_PRIORITY_MAP[priority] || 'Unknown',
    priority:    issue.fields.priority?.name || '',
    state:       issue.fields.status?.name || '',
    areaPath:    (issue.fields.components || []).map(c => c.name).join(', '),
    createdDate: (issue.fields.created || '').slice(0, 10),
    owner:       issue.fields.assignee?.displayName || 'Unassigned',
  };
}

// ─── Provider interface ───────────────────────────────────────────────────────

export async function fetchTestStats({ name, testCycleKey }) {
  if (!testCycleKey) {
    console.warn(`  ⚠  ${name}: no testCycleKey configured, skipping test stats.`);
    return { planned: 0, executed: 0, passed: 0, failed: 0, notStarted: 0, inProgress: 0 };
  }

  let executions;
  try {
    executions = await fetchCycleExecutions(testCycleKey);
  } catch (e) {
    console.warn(`  ⚠  ${name}: Zephyr fetch failed — ${e.message.split('\n')[0]}`);
    return { planned: 0, executed: 0, passed: 0, failed: 0, notStarted: 0, inProgress: 0 };
  }

  console.log(`    → ${name}: ${executions.length} execution(s) in cycle ${testCycleKey}`);
  let planned = 0, executed = 0, passed = 0, failed = 0, notStarted = 0, inProgress = 0;
  for (const ex of executions) {
    planned++;
    const statusKey = (ex.status?.name || 'NOT_EXECUTED').toUpperCase().replace(/ /g, '_');
    const mapped    = ZEPHYR_STATUS_MAP[statusKey] || 'notStarted';
    if      (mapped === 'passed')     { passed++;     executed++; }
    else if (mapped === 'failed')     { failed++;     executed++; }
    else if (mapped === 'inProgress') { inProgress++; executed++; }
    else                              { notStarted++; }
  }
  return { planned, executed, passed, failed, notStarted, inProgress };
}

export async function runQuery({ name }, jql) {
  if (!jql) {
    console.warn(`  ⚠  ${name}: no JQL provided for this query, skipping.`);
    return [];
  }
  if (process.env.DEBUG) console.log(`\n  [DEBUG] ${name} JQL: ${jql}\n`);
  let issues;
  try {
    issues = await jiraSearch(jql, ['summary', 'status', 'priority', 'assignee', 'created', 'components']);
  } catch (e) {
    console.warn(`  ⚠  ${name}: query failed — ${e.message.split('\n')[0]}`);
    return [];
  }
  return issues.map(shapeIssue);
}
