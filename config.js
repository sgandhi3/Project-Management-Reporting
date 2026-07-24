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
// Define any number of data fetches here. Each entry runs once per workstream.
// Results land at d[key][workstreamName] and are accessible in variables.js.
//
// Each query has a section per provider written in that provider's native language:
//   ado  — WIQL config object (workItemType, excludeStates, fields, fieldMap, orderBy)
//   jira — function(ws) returning a JQL string (used by both Jira and Zephyr Scale)
//
// To add a query: duplicate any block, give it a new key, adjust the config.
// To remove a query: delete the block. Nothing else needs to change.

export const QUERIES = [
  {
    key:      'bugs',
    label:    'Open bugs',
    fallback: [],

    ado: {
      workItemType:  'Bug',
      excludeStates: ['Closed', 'Resolved'],
      orderBy:       '[Microsoft.VSTS.Common.Severity] ASC',
      fields:        ADO_FIELDS,
      fieldMap:      ADO_FIELD_MAP,
    },

    jira: ws => `project = "${ws.projectKey}" AND issuetype = Bug AND status != Resolved and status != Closed`,

    // consolidateKey sets the name used in variables.js: d.allBugs
    // consolidate receives { PDM: [...], Benefits: [...], ... } and returns whatever you want
    consolidateKey: 'bugsBySeverityPriority',
    consolidate: allResults => {
      let total = 0, sev1 = 0, sev2 = 0, sev3 = 0, sev4 = 0;
      let pri1  = 0, pri2  = 0, pri3  = 0, pri4  = 0;
      for (const bugs of Object.values(allResults)) {
        for (const b of bugs) {
          total++;
          if      (b.severity === '1 - Critical') sev1++;
          else if (b.severity === '2 - High')     sev2++;
          else if (b.severity === '3 - Medium')   sev3++;
          else if (b.severity === '4 - Low')      sev4++;

          const p = String(b.priority || '').toLowerCase();
          if      (p === '1' || p === 'highest' || p === 'critical') pri1++;
          else if (p === '2' || p === 'high')                        pri2++;
          else if (p === '3' || p === 'medium')                      pri3++;
          else if (p === '4' || p === 'low' || p === 'lowest')       pri4++;
        }
      }
      return {
        total: total,
        severity: { total, sev1, sev2, sev3, sev4 },
        priority: { pri1, pri2, pri3, pri4 },
      };
    },

  },
    {
    key:      'closedBugs',
    label:    'Closed bugs',
    fallback: [],

    ado: {
      workItemType:  'Bug',
      excludeStates: ['Active', 'Resolved', 'New','Blocked'],
      orderBy:       '[Microsoft.VSTS.Common.Severity] ASC',
      fields:        ADO_FIELDS,
      fieldMap:      ADO_FIELD_MAP,
    },

    jira: ws => `project = "${ws.projectKey}" AND issuetype = Bug AND statusCategory != Done ORDER BY priority ASC`,

    // consolidateKey sets the name used in variables.js: d.allBugs
    // consolidate receives { PDM: [...], Benefits: [...], ... } and returns whatever you want
   // consolidateKey: 'bugsBySeverity',
    // consolidate: allResults => {
    //   let total = 0, sev1 = 0, sev2 = 0, sev3 = 0, sev4 = 0;
    //   for (const bugs of Object.values(allResults)) {
    //     for (const b of bugs) {
    //       total++;
    //       if      (b.severity === '1 - Critical') sev1++;
    //       else if (b.severity === '2 - High')     sev2++;
    //       else if (b.severity === '3 - Medium')   sev3++;
    //       else if (b.severity === '4 - Low')      sev4++;
    //     }
    //   }
    //   return { total, sev1, sev2, sev3, sev4 };
    // },
  },

  // ── Example: open tasks with a count-by-month consolidation ───────────────
  // Adding consolidate lets you roll up any query into a summary object that
  // you can then reference in variables.js however you like.
  //
  // {
  //   key:      'openTasks',
  //   label:    'Open tasks',
  //   fallback: [],
  //   ado: {
  //     workItemType:  'Task',
  //     excludeStates: ['Closed', 'Done'],
  //     orderBy:       '[System.CreatedDate] ASC',
  //     fields:        ADO_FIELDS,
  //     fieldMap:      ADO_FIELD_MAP,
  //   },
  //   jira: ws => `project = "${ws.projectKey}" AND issuetype = Task AND statusCategory != Done ORDER BY created ASC`,
  //
  //   // Produces d.tasksByMonth = { '2026-01': 5, '2026-02': 6, '2026-03': 10, ... }
  //   consolidateKey: 'tasksByMonth',
  //   consolidate: allResults => {
  //     const byMonth = {};
  //     for (const item of Object.values(allResults).flat()) {
  //       const month = (item.createdDate || '').slice(0, 7); // "2026-01"
  //       if (month) byMonth[month] = (byMonth[month] || 0) + 1;
  //     }
  //     return byMonth;
  //   },
  // },

  // ── Example: closed bugs ──────────────────────────────────────────────────
  // {
  //   key:      'closedBugs',
  //   label:    'Closed bugs',
  //   fallback: [],
  //   ado: {
  //     workItemType:  'Bug',
  //     excludeStates: [],
  //     includeStates: ['Closed', 'Resolved'],
  //     orderBy:       '[System.CreatedDate] DESC',
  //     fields:        ADO_FIELDS,
  //     fieldMap:      ADO_FIELD_MAP,
  //   },
  //   jira: ws => `project = "${ws.projectKey}" AND issuetype = Bug AND statusCategory = Done`,
  //
  //   consolidateKey: 'closedBugsByMonth',
  //   consolidate: allResults => {
  //     const byMonth = {};
  //     for (const item of Object.values(allResults).flat()) {
  //       const month = (item.createdDate || '').slice(0, 7);
  //       if (month) byMonth[month] = (byMonth[month] || 0) + 1;
  //     }
  //     return byMonth;
  //   },
  // },
];
