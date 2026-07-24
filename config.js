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
  'blocked':     'notStarted',
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
  'BLOCKED':      'notStarted',
};

// ─── Workstreams ───────────────────────────────────────────────────────────────
//
// One entry per workstream. The name is used as the key everywhere — in the data
// object, in variables.js (d.stats.PDM), and in the snapshot printout.
//
// Fill in only the fields for the provider you're using (set in TEST_PROVIDER in .env).
// Fields for other providers are ignored.

export const WORKSTREAMS = [
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
//   Works for any field on any query — severity, priority, assignee, state, etc.
//
// Provider sections:
//   ado  — structured config OR { wiql: string } for raw WIQL
//   jira — function(ws) returning a JQL string

export const QUERIES = [
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

  // ── Example: project-wide query (scope: 'global') ──────────────────────────
  // Runs once, no area-path filter, returns all items across the entire project.
  // Useful when you want a single de-duplicated list rather than per-workstream counts.
  //
  // {
  //   key:           'allBugs',
  //   label:         'All open bugs (project-wide)',
  //   fallback:      [],
  //   scope:         'global',
  //   groupByFields: ['severity', 'owner'],
  //   ado: {
  //     workItemType:  'Bug',
  //     excludeStates: ['Closed', 'Resolved'],
  //     orderBy:       '[Microsoft.VSTS.Common.Severity] ASC',
  //     fields:        ADO_FIELDS,
  //     fieldMap:      ADO_FIELD_MAP,
  //   },
  //   jira: () => `issuetype = Bug AND status != Resolved AND status != Closed`,
  // },

  // ── Example: group by assignee ────────────────────────────────────────────
  // {
  //   key:           'userStories',
  //   label:         'Open user stories',
  //   fallback:      [],
  //   scope:         'workstream',
  //   groupByFields: ['owner', 'state'],
  //   ado: {
  //     workItemType:  'User Story',
  //     excludeStates: ['Closed', 'Resolved'],
  //     orderBy:       '[System.CreatedDate] ASC',
  //     fields:        ADO_FIELDS,
  //     fieldMap:      ADO_FIELD_MAP,
  //   },
  //   jira: ws => `project = "${ws.projectKey}" AND issuetype = Story AND statusCategory != Done`,
  // },
];
