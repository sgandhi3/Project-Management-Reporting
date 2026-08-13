import { readFileSync, existsSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const UI_CONFIG_PATH = path.join(__dirname, 'ui-config.json');
const uiConfig = existsSync(UI_CONFIG_PATH) ? JSON.parse(readFileSync(UI_CONFIG_PATH, 'utf8')) : null;

// ─── ADO field definitions ─────────────────────────────────────────────────────
//
// Shared across all ADO query blocks. Add a field here and expose it via
// ADO_FIELD_MAP — it then becomes accessible in variables.js on every result item.

export const ADO_FIELDS = [
  '[System.Id]',
  '[System.Title]',
  '[System.State]',
  '[System.AreaPath]',
  '[System.AssignedTo]',
  '[System.CreatedDate]',
  '[Microsoft.VSTS.Common.Severity]',
  '[Microsoft.VSTS.Common.Priority]',
];

export const ADO_FIELD_MAP = {
  id:          wi => wi.id,
  title:       wi => wi.fields['System.Title']                      || '',
  severity:    wi => wi.fields['Microsoft.VSTS.Common.Severity']    || 'Unknown',
  priority:    wi => wi.fields['Microsoft.VSTS.Common.Priority']    || '',
  state:       wi => wi.fields['System.State']                      || '',
  areaPath:    wi => wi.fields['System.AreaPath']                   || '',
  createdDate: wi => (wi.fields['System.CreatedDate']               || '').slice(0, 10),
  owner:       wi => (wi.fields['System.AssignedTo']                || {}).displayName || 'Unassigned',
};

// ─── Status / severity maps ────────────────────────────────────────────────────
//
// Jira and Zephyr use text names — these translate them to the standard values
// the report expects. Add entries here for any custom statuses your project uses.

export const JIRA_STATUS_MAP = {
  'pass':        'passed',
  'passed':      'passed',
  'fail':        'failed',
  'failed':      'failed',
  'in progress': 'inProgress',
  'in-progress': 'inProgress',
  'to do':       'notStarted',
  'backlog':     'notStarted',
  'not started': 'notStarted',
  'blocked':     'blocked',
  'paused':      'paused',
};

export const JIRA_PRIORITY_MAP = {
  'highest':  '1 - Critical',
  'critical': '1 - Critical',
  'high':     '2 - High',
  'medium':   '3 - Medium',
  'low':      '4 - Low',
  'lowest':   '4 - Low',
};

export const ZEPHYR_STATUS_MAP = {
  'PASS':         'passed',
  'FAIL':         'failed',
  'IN_PROGRESS':  'inProgress',
  'NOT_EXECUTED': 'notStarted',
  'UNEXECUTED':   'notStarted',
  'BLOCKED':      'blocked',
  'PAUSED':       'paused',
};

// ─── Workstreams ───────────────────────────────────────────────────────────────
//
// One entry per workstream. The name is used as the key everywhere — in the data
// object, in variables.js (d.stats.PDM), and in the snapshot printout.
//
// When ui-config.json exists its workstreams take precedence; otherwise the
// env-var fallback below is used.

export const WORKSTREAMS = uiConfig?.workstreams?.length ? uiConfig.workstreams : [
  {
    name: 'PDM',
    // ADO
    planId:       process.env.PLAN_ID,
    sitSuiteId:   process.env.SIT_SUITE_PDM,
    areaPath:     process.env.PDMAreaPath,
    // Jira / Zephyr
    projectKey:   process.env.JIRA_PROJECT_PDM,
    testCycleKey: process.env.ZEPHYR_CYCLE_PDM,
  },
  {
    name: 'Benefits',
    // ADO
    planId:       process.env.PLAN_ID,
    sitSuiteId:   process.env.SIT_SUITE_BENEFITS,
    areaPath:     process.env.BenefitsAreaPath,
    // Jira / Zephyr
    projectKey:   process.env.JIRA_PROJECT_BENEFITS,
    testCycleKey: process.env.ZEPHYR_CYCLE_BENEFITS,
  },
  {
    name: 'Enrollment',
    // ADO
    planId:       process.env.PLAN_ID,
    sitSuiteId:   process.env.SIT_SUITE_ENROLLMENT,
    areaPath:     process.env.EnrollmentAreaPath,
    // Jira / Zephyr
    projectKey:   process.env.JIRA_PROJECT_ENROLLMENT,
    testCycleKey: process.env.ZEPHYR_CYCLE_ENROLLMENT,
  },
  {
    name: 'EDI',
    // ADO
    planId:       process.env.PLAN_ID,
    sitSuiteId:   process.env.SIT_SUITE_EDI,
    areaPath:     process.env.EDIAreaPath,
    // Jira / Zephyr
    projectKey:   process.env.JIRA_PROJECT_EDI,
    testCycleKey: process.env.ZEPHYR_CYCLE_EDI,
  },
];

// ─── Queries ───────────────────────────────────────────────────────────────────
//
// Each entry defines one data fetch. Two modes (set via `scope`):
//
//   scope: 'workstream' (default)
//     Runs once per workstream, filtered by that workstream's area path.
//     Results: d[key] = { PDM: [...], Benefits: [...], ... }
//              d[key+'Total'] = total item count across all workstreams
//
//   scope: 'global'
//     Runs ONCE for the whole project with no area-path filter.
//     Results: d[key] = [...flat array of all items...]
//
// groupByFields: ['severity', 'priority']
//   Automatically produces d[key+'BySeverity'], d[key+'ByPriority'], etc.
//   Each groupBy result: { total: N, 'value1': count, 'value2': count, ... }
//
// When ui-config.json exists its queries take precedence via buildQueriesFromConfig.

function buildQueriesFromConfig(queryConfigs) {
  return queryConfigs
    .filter(q => q.enabled !== false)
    .map(q => {
      const groupByFields = Array.isArray(q.groupByFields)
        ? q.groupByFields
        : (q.groupByFields ? String(q.groupByFields).split(',').map(s => s.trim()).filter(Boolean) : []);

      return {
        key:           q.key,
        label:         q.label,
        fallback:      [],
        scope:         q.scope || 'workstream',
        groupByFields,
        ...(q.wiqlTemplate ? {
          ado: ws => ({
            wiql: q.wiqlTemplate.replace(/\{\{(\w+)\}\}/g, (_, k) => (ws && ws[k]) || process.env[k] || ''),
            fieldMap: ADO_FIELD_MAP,
          }),
        } : q.ado ? {
          ado: {
            workItemType:  q.ado.workItemType,
            excludeStates: q.ado.excludeStates || [],
            includeStates: q.ado.includeStates || [],
            orderBy:       q.ado.orderBy || '[System.Id] ASC',
            fields:        ADO_FIELDS,
            fieldMap:      ADO_FIELD_MAP,
          },
        } : {}),
        ...(q.jiraTemplate ? {
          jira: ws => q.jiraTemplate.replace(/\{\{(\w+)\}\}/g, (_, k) => (ws && ws[k]) || process.env[k] || ''),
        } : {}),
      };
    });
}

export const SETTINGS = uiConfig?.settings || {};

export const QUERIES = uiConfig?.queries?.length
  ? buildQueriesFromConfig(uiConfig.queries)
  : [
    {
      key:           'bugs',
      label:         'Open bugs',
      fallback:      [],
      scope:         'workstream',
      groupByFields: ['severity', 'priority'],

      ado: {
        workItemType:  'Bug',
        excludeStates: ['Closed', 'Resolved'],
        orderBy:       '[Microsoft.VSTS.Common.Severity] ASC',
        fields:        ADO_FIELDS,
        fieldMap:      ADO_FIELD_MAP,
      },

      jira: ws => `project = "${ws.projectKey}" AND issuetype = Bug AND status != Resolved and status != Closed`,
    },

    {
      key:           'closedBugs',
      label:         'Closed bugs',
      fallback:      [],
      scope:         'workstream',
      groupByFields: [],

      ado: {
        workItemType:  'Bug',
        excludeStates: ['Active', 'Resolved', 'New', 'Blocked'],
        orderBy:       '[Microsoft.VSTS.Common.Severity] ASC',
        fields:        ADO_FIELDS,
        fieldMap:      ADO_FIELD_MAP,
      },

      jira: ws => `project = "${ws.projectKey}" AND issuetype = Bug AND statusCategory = Done ORDER BY priority ASC`,
    },
  ];
