import './env.js';  // must be first — loads .env before config.js reads process.env
import { WORKSTREAMS, QUERIES } from './config.js';

// ─── Provider selection ───────────────────────────────────────────────────────

// Set TEST_PROVIDER in .env to switch between data sources.
// Supported values: ado | jira
const PROVIDER = (process.env.TEST_PROVIDER || 'ado').toLowerCase();

// Set OUTPUT_FORMAT in .env — comma-separated to run multiple extensions.
// e.g. OUTPUT_FORMAT=ppt,ai-summary   runs both
// Add more by creating extensions/<name>.js
const OUTPUTS = (process.env.OUTPUT_FORMAT || 'ppt').toLowerCase().split(',').map(s => s.trim()).filter(Boolean);

// ─── Data aggregation ─────────────────────────────────────────────────────────

function sumStats(statsMap) {
  const totals = { planned: 0, executed: 0, passed: 0, failed: 0, notStarted: 0, inProgress: 0 };
  for (const s of Object.values(statsMap)) {
    for (const k of Object.keys(totals)) totals[k] += s[k];
  }
  return totals;
}

// ─── Orchestration ────────────────────────────────────────────────────────────

async function collectAllData(provider) {
  console.log(`\nFetching data via ${PROVIDER} provider...\n`);

  const data = { stats: {} };
  for (const q of QUERIES) data[q.key] = {};

  for (const ws of WORKSTREAMS) {
    console.log(`${ws.name}:`);

    process.stdout.write(`  Test execution stats... `);
    try {
      data.stats[ws.name] = await provider.fetchTestStats(ws);
      const s = data.stats[ws.name];
      console.log(`planned=${s.planned} executed=${s.executed} passed=${s.passed} failed=${s.failed}`);
    } catch (e) {
      console.log(`failed — ${e.message.split('\n')[0]}`);
      data.stats[ws.name] = { planned: 0, executed: 0, passed: 0, failed: 0, notStarted: 0, inProgress: 0 };
    }

    for (const q of QUERIES) {
      process.stdout.write(`  ${q.label}... `);
      try {
        const providerConfig = q[PROVIDER];
        const config = typeof providerConfig === 'function' ? providerConfig(ws) : providerConfig;
        const result = await provider.runQuery(ws, config);
        data[q.key][ws.name] = result;
        console.log(Array.isArray(result) ? `${result.length} found` : JSON.stringify(result));
      } catch (e) {
        console.log(`failed — ${e.message.split('\n')[0]}`);
        data[q.key][ws.name] = q.fallback;
      }
    }

    console.log('');
  }

  data.consolidatedData = sumStats(data.stats);
  console.log('Overall Execution:'
    + '\n  Planned:     ' + data.consolidatedData.planned
    + '\n  Executed:    ' + data.consolidatedData.executed
    + '\n  Passed:      ' + data.consolidatedData.passed
    + '\n  Failed:      ' + data.consolidatedData.failed
    + '\n  Not Started: ' + data.consolidatedData.notStarted
    + '\n  In Progress: ' + data.consolidatedData.inProgress + '\n');

  for (const q of QUERIES) {
    if (!q.consolidate) continue;
    const key = q.consolidateKey || q.key + 'Summary';
    data[key] = q.consolidate(data[q.key]);
    console.log(`${key}:`, JSON.stringify(data[key]));
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
    passed: s.passed, failed: s.failed, notStarted: s.notStarted, inProgress: s.inProgress,
  }));
  console.table(statsRows);

  const c = data.consolidatedData;
  console.log(' d.consolidatedData  (all workstreams combined)');
  console.table([{ planned: c.planned, executed: c.executed, passed: c.passed, failed: c.failed, notStarted: c.notStarted, inProgress: c.inProgress }]);

  for (const q of QUERIES) {
    const results = data[q.key] || {};
    console.log(` d.${q.key}  (${q.label})\n`);
    const rows = Object.entries(results).map(([ws, items]) => {
      const row = { workstream: ws, count: Array.isArray(items) ? items.length : '—' };
      const sample = Array.isArray(items) ? items[0] : items;
      if (sample) row['available fields'] = Object.keys(sample).join(', ');
      return row;
    });
    console.table(rows);

    const firstWs    = Object.values(results).find(arr => Array.isArray(arr) && arr.length > 0);
    const sampleItem = firstWs?.[0];
    if (sampleItem) {
      console.log(`  sample item → d.${q.key}.<workstream>[0]:`);
      console.table([sampleItem]);
    }
  }

  for (const q of QUERIES) {
    if (!q.consolidate) continue;
    const key = q.consolidateKey || q.key + 'Summary';
    const val = data[key];
    if (!val) continue;
    const hasNested = Object.values(val).some(v => v !== null && typeof v === 'object' && !Array.isArray(v));
    if (hasNested) {
      for (const [subKey, subVal] of Object.entries(val)) {
        if (subVal !== null && typeof subVal === 'object' && !Array.isArray(subVal)) {
          console.log(` d.${key}.${subKey}  (${q.label} — by ${subKey})\n`);
          console.table([subVal]);
        }
      }
    } else {
      console.log(` d.${key}  (${q.label} — consolidated)\n`);
      console.table([val]);
    }
  }

  console.log(divider + '\n');
}

// ─── Entry point ──────────────────────────────────────────────────────────────

async function main() {
  const provider  = await import(`./providers/${PROVIDER}.js`);
  const data      = await collectAllData(provider);
  printDataSnapshot(data);

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
