import './env.js';  // must be first — loads .env before config.js reads process.env
import { WORKSTREAMS, QUERIES, SETTINGS } from './config.js';

const PROVIDER = (process.env.TEST_PROVIDER || 'ado').toLowerCase();

// ui-config.json outputFormats (set by the UI checkboxes) takes precedence over OUTPUT_FORMAT in .env
// so that unchecking a format in the UI immediately takes effect without editing .env manually
const OUTPUTS = SETTINGS.outputFormats?.length
  ? SETTINGS.outputFormats
  : (process.env.OUTPUT_FORMAT || 'ppt').toLowerCase().split(',').map(s => s.trim()).filter(Boolean);

// ─── Utilities ────────────────────────────────────────────────────────────────

function sumStats(statsMap) {
  const totals = { planned: 0, executed: 0, passed: 0, failed: 0, notStarted: 0, inProgress: 0, blocked: 0, paused: 0 };
  for (const s of Object.values(statsMap)) {
    for (const k of Object.keys(totals)) totals[k] += (s[k] ?? 0);
  }
  return totals;
}

// Groups an array of items by a field value, returning { total, value: count, ... }
function groupBy(items, field) {
  const result = { total: items.length };
  for (const item of items) {
    const val = String(item[field] ?? '(none)');
    result[val] = (result[val] || 0) + 1;
  }
  return result;
}

// ─── Orchestration ────────────────────────────────────────────────────────────

async function collectAllData(provider) {
  console.log(`\nFetching data via ${PROVIDER} provider...\n`);

  const data = { stats: {}, subStats: {} };
  for (const q of QUERIES) {
    data[q.key] = q.scope === 'global' ? [] : {};
  }

  // ── Global queries — run once for the whole project, no area-path filter ────
  const globalQueries = QUERIES.filter(q => q.scope === 'global');
  if (globalQueries.length) {
    console.log('Global queries:');
    for (const q of globalQueries) {
      process.stdout.write(`  ${q.label} (all workstreams)... `);
      try {
        const providerConfig = q[PROVIDER];
        const config = typeof providerConfig === 'function' ? providerConfig({}) : providerConfig;
        const result = await provider.runQuery({ name: 'global', areaPath: '' }, config);
        data[q.key] = result;
        console.log(`${result.length} found`);
      } catch (e) {
        console.log(`failed — ${e.message.split('\n')[0]}`);
        data[q.key] = q.fallback ?? [];
      }
    }
    console.log('');
  }

  // ── Per-workstream queries — run once per workstream ─────────────────────────
  const wsQueries = QUERIES.filter(q => q.scope !== 'global');

  for (const ws of WORKSTREAMS) {
    console.log(`${ws.name}:`);

    process.stdout.write(`  Test execution stats... `);
    try {
      data.stats[ws.name] = await provider.fetchTestStats(ws);
      const s = data.stats[ws.name];
      console.log(`planned=${s.planned} executed=${s.executed} passed=${s.passed} failed=${s.failed}`);
    } catch (e) {
      console.log(`failed — ${e.message.split('\n')[0]}`);
      data.stats[ws.name] = { planned: 0, executed: 0, passed: 0, failed: 0, notStarted: 0, inProgress: 0, blocked: 0 };
    }

    if (provider.fetchSubSuiteStats) {
      process.stdout.write(`  Sub-suite breakdown... `);
      try {
        data.subStats[ws.name] = await provider.fetchSubSuiteStats(ws);
        const count = Object.keys(data.subStats[ws.name]).length;
        console.log(count ? Object.keys(data.subStats[ws.name]).join(', ') : 'no sub-suites');
      } catch (e) {
        console.log(`failed — ${e.message.split('\n')[0]}`);
        data.subStats[ws.name] = {};
      }
    }

    for (const q of wsQueries) {
      process.stdout.write(`  ${q.label}... `);
      try {
        const providerConfig = q[PROVIDER];
        const config = typeof providerConfig === 'function' ? providerConfig(ws) : providerConfig;
        const result = await provider.runQuery(ws, config);
        data[q.key][ws.name] = result;
        console.log(Array.isArray(result) ? `${result.length} found` : JSON.stringify(result));
      } catch (e) {
        console.log(`failed — ${e.message.split('\n')[0]}`);
        data[q.key][ws.name] = q.fallback ?? [];
      }
    }

    console.log('');
  }

  // ── Consolidated test stats ───────────────────────────────────────────────────
  data.consolidatedData = sumStats(data.stats);
  console.log('Overall Execution:'
    + '\n  Planned:     ' + data.consolidatedData.planned
    + '\n  Executed:    ' + data.consolidatedData.executed
    + '\n  Passed:      ' + data.consolidatedData.passed
    + '\n  Failed:      ' + data.consolidatedData.failed
    + '\n  Not Started: ' + data.consolidatedData.notStarted
    + '\n  In Progress: ' + data.consolidatedData.inProgress
    + '\n  Blocked:     ' + data.consolidatedData.blocked
    + '\n  Paused:      ' + data.consolidatedData.paused + '\n');

  // ── Auto-total for every per-workstream query (d.bugsTotal, etc.) ─────────────
  for (const q of wsQueries) {
    data[`${q.key}Total`] = Object.values(data[q.key])
      .reduce((s, a) => s + (Array.isArray(a) ? a.length : 0), 0);
  }

  // ── Generic group-by aggregations ─────────────────────────────────────────────
  // For each field in groupByFields, produce d.<key>By<Field> = { total, value: count, ... }
  for (const q of QUERIES) {
    const fields = q.groupByFields || [];
    if (!fields.length) continue;

    const allItems = q.scope === 'global'
      ? data[q.key]
      : Object.values(data[q.key]).flat();

    for (const field of fields) {
      const cap = field.charAt(0).toUpperCase() + field.slice(1);
      const key = `${q.key}By${cap}`;
      data[key] = groupBy(allItems, field);
      console.log(`${key}: total=${data[key].total}`);
    }
  }

  return data;
}

// ─── Data snapshot ────────────────────────────────────────────────────────────

function printDataSnapshot(data) {
  const divider = '─'.repeat(64);
  console.log('\n' + divider);
  console.log(' DATA SNAPSHOT  — everything accessible as d.* in variables.js');
  console.log(divider);

  console.log('\n d.stats  (test execution per workstream)\n');
  const statsRows = Object.entries(data.stats).map(([ws, s]) => ({
    workstream: ws, planned: s.planned, executed: s.executed,
    passed: s.passed, failed: s.failed, notStarted: s.notStarted, inProgress: s.inProgress, blocked: s.blocked ?? 0,
  }));
  console.table(statsRows);

  const c = data.consolidatedData;
  console.log(' d.consolidatedData  (all workstreams combined)');
  console.table([{ planned: c.planned, executed: c.executed, passed: c.passed, failed: c.failed, notStarted: c.notStarted, inProgress: c.inProgress, blocked: c.blocked ?? 0 }]);

  const subStatEntries = Object.entries(data.subStats).filter(([, v]) => Object.keys(v).length > 0);
  if (subStatEntries.length) {
    console.log(' d.subStats  (per-workstream sub-suite breakdown)\n');
    for (const [wsName, suites] of subStatEntries) {
      console.log(`  ${wsName}:`);
      const rows = Object.entries(suites).map(([suite, s]) => ({ suite, planned: s.planned, executed: s.executed, passed: s.passed, failed: s.failed, blocked: s.blocked ?? 0 }));
      console.table(rows);
    }
  }

  for (const q of QUERIES) {
    const results = data[q.key];

    if (q.scope === 'global') {
      const items = Array.isArray(results) ? results : [];
      console.log(` d.${q.key}  (${q.label} — global)   total: ${items.length}\n`);
      if (items[0]) {
        console.log(`  available fields: ${Object.keys(items[0]).join(', ')}`);
        console.table([items[0]]);
      }
    } else {
      console.log(` d.${q.key}  (${q.label})   d.${q.key}Total = ${data[`${q.key}Total`]}\n`);
      const rows = Object.entries(results || {}).map(([ws, items]) => {
        const row = { workstream: ws, count: Array.isArray(items) ? items.length : '—' };
        const sample = Array.isArray(items) ? items[0] : items;
        if (sample) row['available fields'] = Object.keys(sample).join(', ');
        return row;
      });
      console.table(rows);

      const firstWs    = Object.values(results || {}).find(arr => Array.isArray(arr) && arr.length > 0);
      const sampleItem = firstWs?.[0];
      if (sampleItem) {
        console.log(`  sample item → d.${q.key}.<workstream>[0]:`);
        console.table([sampleItem]);
      }
    }
  }

  // GroupBy result tables
  for (const q of QUERIES) {
    for (const field of q.groupByFields || []) {
      const cap = field.charAt(0).toUpperCase() + field.slice(1);
      const key = `${q.key}By${cap}`;
      if (data[key]) {
        console.log(` d.${key}  (${q.label} — by ${field})\n`);
        console.table([data[key]]);
      }
    }
  }

  console.log(divider + '\n');
}

// ─── Entry point ──────────────────────────────────────────────────────────────

const args = process.argv.slice(2);

async function main() {
  const provider  = await import(`./providers/${PROVIDER}.js`);
  const data      = await collectAllData(provider);

  // --preview: output data as JSON for the UI Data tab, skip extensions
  if (args.includes('--preview')) {
    process.stdout.write('__PREVIEW__' + JSON.stringify(data, (_k, v) =>
      // Drop raw item arrays to keep payload small — keep counts/groupBy/stats
      Array.isArray(v) && v.length > 0 && typeof v[0] === 'object' && !Array.isArray(v[0]) ? `[${v.length} items]` : v
    ) + '\n');
    return;
  }

  printDataSnapshot(data);

  // Run AI extensions first — they populate data._aiSummary / data._aiNarratives
  // for use by ppt/excel. ai-narrative runs whenever ppt output is requested,
  // since the template's {{AI_...}} tokens live in the ppt deck.
  const aiIdx = OUTPUTS.indexOf('ai-summary');
  if (aiIdx !== -1) {
    const aiExt = await import('./extensions/ai-summary.js');
    await aiExt.generate(data);
    OUTPUTS.splice(aiIdx, 1); // prevent running again below
  }

  if (OUTPUTS.includes('ppt')) {
    const narrativeExt = await import('./extensions/ai-narrative.js');
    await narrativeExt.generate(data);
  }

  // Then run remaining extensions (ppt, excel, sharepoint, etc.)
  for (const output of OUTPUTS) {
    let extension;
    try {
      extension = await import(`./extensions/${output}.js`);
    } catch {
      console.error(`\n❌  Unknown OUTPUT_FORMAT "${output}" — no extensions/${output}.js found.`);
      process.exit(1);
    }
    await extension.generate(data);
  }
}

main().catch(err => { console.error('\n❌  Fatal error:', err.message.split('\n')[0]); process.exit(1); });
