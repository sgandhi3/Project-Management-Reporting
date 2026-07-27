// Web UI server for MMO Report Generator
//
// Starts an Express server on port 3001 (override with UI_PORT env var).
// Serves a single-page app at / and exposes a REST API for reading and
// writing config.js (via ui-config.json), .env, and running gather-data.js.
//
// Usage: node extensions/ui.js   (or: npm run ui)

import express from 'express';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { spawn } from 'child_process';
import cron from 'node-cron';

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);
const projectRoot = path.join(__dirname, '..');

const PORT = parseInt(process.env.UI_PORT || '3001', 10);
const UI_CONFIG_PATH = path.join(projectRoot, 'ui-config.json');

// Sensitive key pattern: mask values whose key matches
const SENSITIVE_RE = /TOKEN|PAT|KEY|SECRET|PASSWORD/i;

// ─── Scheduler ────────────────────────────────────────────────────────────────

let scheduledTask = null;

function applySchedule(settings) {
  if (scheduledTask) { scheduledTask.stop(); scheduledTask = null; }
  const s = settings?.schedule;
  if (!s?.enabled || !s?.cron) return;
  if (!cron.validate(s.cron)) {
    console.warn(`[Scheduler] Invalid cron expression: "${s.cron}"`);
    return;
  }
  scheduledTask = cron.schedule(s.cron, () => {
    console.log(`[Scheduler] Triggering scheduled run (${s.cron})...`);
    triggerRun();
  });
  console.log(`[Scheduler] Active — cron: ${s.cron}`);
}

// ─── Default ui-config (written on first GET /api/config) ────────────────────

const DEFAULT_UI_CONFIG = {
  settings: {
    fetchTestStats: true,
    outputFormats: ['ppt'],
    schedule: { enabled: false, cron: '0 9 * * 1' },
  },
  workstreams: [
    { name: 'PDM',        planId: '', sitSuiteId: '', areaPath: '', projectKey: '', testCycleKey: '' },
    { name: 'Benefits',   planId: '', sitSuiteId: '', areaPath: '', projectKey: '', testCycleKey: '' },
    { name: 'Enrollment', planId: '', sitSuiteId: '', areaPath: '', projectKey: '', testCycleKey: '' },
    { name: 'EDI',        planId: '', sitSuiteId: '', areaPath: '', projectKey: '', testCycleKey: '' },
  ],
  queries: [
    {
      key: 'bugs', label: 'Open bugs', enabled: true,
      scope: 'workstream', groupByFields: ['severity', 'priority'],
      wiqlTemplate: "SELECT [System.Id], [System.Title], [System.State], [System.AreaPath], [System.AssignedTo], [System.CreatedDate], [Microsoft.VSTS.Common.Severity], [Microsoft.VSTS.Common.Priority] FROM WorkItems WHERE [System.WorkItemType] = 'Bug' AND [System.TeamProject] = '{{ADO_PROJECT}}' AND [System.State] <> 'Closed' AND [System.State] <> 'Resolved' AND [System.AreaPath] UNDER '{{areaPath}}' ORDER BY [Microsoft.VSTS.Common.Severity] ASC",
      jiraTemplate: 'project = "{{projectKey}}" AND issuetype = Bug AND status != Resolved and status != Closed',
    },
    {
      key: 'closedBugs', label: 'Closed bugs', enabled: true,
      scope: 'workstream', groupByFields: [],
      wiqlTemplate: "SELECT [System.Id], [System.Title], [System.State], [System.AreaPath], [System.AssignedTo], [System.CreatedDate], [Microsoft.VSTS.Common.Severity], [Microsoft.VSTS.Common.Priority] FROM WorkItems WHERE [System.WorkItemType] = 'Bug' AND [System.TeamProject] = '{{ADO_PROJECT}}' AND [System.State] NOT IN ('Active', 'Resolved', 'New', 'Blocked') AND [System.AreaPath] UNDER '{{areaPath}}' ORDER BY [Microsoft.VSTS.Common.Severity] ASC",
      jiraTemplate: 'project = "{{projectKey}}" AND issuetype = Bug AND statusCategory = Done ORDER BY priority ASC',
    },
  ],
};

// ─── .env helpers ─────────────────────────────────────────────────────────────

function readEnvFile() {
  const envPath = path.join(projectRoot, '.env');
  if (!fs.existsSync(envPath)) return {};
  const lines = fs.readFileSync(envPath, 'utf8').split('\n');
  const result = {};
  for (const line of lines) {
    const match = line.match(/^([^#=\s][^=]*)=(.*)$/);
    if (match) result[match[1].trim()] = match[2].trim();
  }
  return result;
}

function writeEnvFile(updates) {
  const envPath = path.join(projectRoot, '.env');
  let content = fs.existsSync(envPath) ? fs.readFileSync(envPath, 'utf8') : '';
  for (const [key, value] of Object.entries(updates)) {
    const regex = new RegExp(`^(${key}=).*$`, 'm');
    if (regex.test(content)) {
      content = content.replace(regex, `$1${value}`);
    } else {
      content = content.trimEnd() + `\n${key}=${value}\n`;
    }
  }
  fs.writeFileSync(envPath, content);
}

// ─── Run state (SSE streaming) ────────────────────────────────────────────────

let currentProcess = null;
let logBuffer      = [];
const MAX_BUFFER   = 500;
const sseClients   = new Set();

function sendToClients(line) {
  logBuffer.push(line);
  if (logBuffer.length > MAX_BUFFER) logBuffer.shift();
  for (const res of sseClients) {
    res.write(`data: ${JSON.stringify(line)}\n\n`);
  }
}

// ─── Express app ─────────────────────────────────────────────────────────────

const app = express();
app.use(express.json({ limit: '1mb' }));

// Serve the single-page UI
app.get('/', (_req, res) => {
  res.sendFile(path.join(projectRoot, 'ui', 'index.html'));
});

// ─── Config API ───────────────────────────────────────────────────────────────

app.get('/api/config', (_req, res) => {
  if (!fs.existsSync(UI_CONFIG_PATH)) {
    fs.writeFileSync(UI_CONFIG_PATH, JSON.stringify(DEFAULT_UI_CONFIG, null, 2));
  }
  try {
    res.json(JSON.parse(fs.readFileSync(UI_CONFIG_PATH, 'utf8')));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.put('/api/config', (req, res) => {
  try {
    fs.writeFileSync(UI_CONFIG_PATH, JSON.stringify(req.body, null, 2));
    applySchedule(req.body.settings);  // restart scheduler if schedule changed
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── Env API ─────────────────────────────────────────────────────────────────

app.get('/api/env', (_req, res) => {
  const env    = readEnvFile();
  const masked = {};
  for (const [k, v] of Object.entries(env)) {
    masked[k] = SENSITIVE_RE.test(k) ? '***' : v;
  }
  res.json(masked);
});

app.put('/api/env', (req, res) => {
  try {
    const { key, value, updates } = req.body;
    const toUpdate = updates || (key !== undefined ? { [key]: value } : {});
    writeEnvFile(toUpdate);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── Run API ─────────────────────────────────────────────────────────────────

function triggerRun() {
  if (currentProcess) return null;  // already running
  logBuffer = [];
  const childEnv = { ...process.env, ...readEnvFile() };
  currentProcess = spawn('node', ['gather-data.js'], { cwd: projectRoot, env: childEnv });
  const handleData = chunk => {
    for (const line of chunk.toString().split('\n')) {
      if (line) sendToClients(line);
    }
  };
  currentProcess.stdout.on('data', handleData);
  currentProcess.stderr.on('data', handleData);
  currentProcess.on('close', code => {
    sendToClients(`[Process exited with code ${code}]`);
    currentProcess = null;
  });
  currentProcess.on('error', err => {
    sendToClients(`[Spawn error: ${err.message}]`);
    currentProcess = null;
  });
  return currentProcess.pid;
}

app.post('/api/run', (_req, res) => {
  if (currentProcess) return res.status(409).json({ error: 'A run is already in progress' });
  const pid = triggerRun();
  res.json({ pid });
});

app.get('/api/run/stream', (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();

  // Replay buffered history to late-connecting clients
  for (const line of logBuffer) {
    res.write(`data: ${JSON.stringify(line)}\n\n`);
  }

  sseClients.add(res);
  req.on('close', () => sseClients.delete(res));
});

app.get('/api/run/status', (_req, res) => {
  res.json({ running: currentProcess !== null });
});

// ─── Variables API ────────────────────────────────────────────────────────────

app.get('/api/variables', (_req, res) => {
  try {
    const uiConfig = fs.existsSync(UI_CONFIG_PATH)
      ? JSON.parse(fs.readFileSync(UI_CONFIG_PATH, 'utf8'))
      : {};
    res.json(uiConfig.variableMappings || []);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.put('/api/variables', (req, res) => {
  try {
    const { variableMappings } = req.body;
    const uiConfig = fs.existsSync(UI_CONFIG_PATH)
      ? JSON.parse(fs.readFileSync(UI_CONFIG_PATH, 'utf8'))
      : {};
    uiConfig.variableMappings = variableMappings;
    fs.writeFileSync(UI_CONFIG_PATH, JSON.stringify(uiConfig, null, 2));
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── Data Paths API ───────────────────────────────────────────────────────────

function buildDataPaths(uiConfig) {
  const paths = [];
  const workstreams = uiConfig.workstreams || [];
  const queries     = uiConfig.queries || [];

  // Test stats per workstream
  const statFields = ['planned', 'executed', 'passed', 'failed', 'notStarted', 'inProgress'];
  for (const ws of workstreams) {
    for (const f of statFields) paths.push(`d.stats.${ws.name}.${f}`);
  }

  // Consolidated stats + calculated percentages
  for (const f of statFields) paths.push(`d.consolidatedData.${f}`);
  paths.push(`d.consolidatedData.executed ? Math.round((d.consolidatedData.passed / d.consolidatedData.executed) * 100) : 0`);
  paths.push(`d.consolidatedData.executed ? Math.round((d.consolidatedData.failed / d.consolidatedData.executed) * 100) : 0`);

  for (const q of queries.filter(q => q.enabled !== false)) {
    const scope        = q.scope || 'workstream';
    const groupByFields = Array.isArray(q.groupByFields)
      ? q.groupByFields
      : (q.groupByFields ? String(q.groupByFields).split(',').map(s => s.trim()).filter(Boolean) : []);

    if (scope === 'global') {
      // Global query — flat array
      paths.push(`d.${q.key}.length`);
    } else {
      // Per-workstream query — per-ws counts + auto-total
      for (const ws of workstreams) {
        paths.push(`d.${q.key}.${ws.name}.length`);
      }
      paths.push(`d.${q.key}Total`);
    }

    // GroupBy paths — one section per field, with concrete values for known fields
    const KNOWN_VALUES = {
      severity: ['1 - Critical', '2 - High', '3 - Medium', '4 - Low', 'Unknown'],
      priority: [1, 2, 3, 4],
      state:    ['Active', 'New', 'Resolved', 'Closed', 'Blocked', 'In Progress'],
      owner:    [],
      assignee: [],
    };
    for (const field of groupByFields) {
      const cap    = field.charAt(0).toUpperCase() + field.slice(1);
      const byKey  = `${q.key}By${cap}`;
      const values = KNOWN_VALUES[field.toLowerCase()] || [];
      paths.push(`d.${byKey}.total`);
      for (const v of values) {
        paths.push(typeof v === 'string' ? `d.${byKey}['${v}']` : `d.${byKey}[${v}]`);
      }
      if (!values.length) {
        paths.push(`d.${byKey}['<run report to see values>']`);
      }
    }
  }

  // Date utility
  paths.push(`new Date().toLocaleDateString('en-US', { timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: 'numeric' })`);

  return paths;
}

app.get('/api/data-paths', (_req, res) => {
  try {
    const uiConfig = fs.existsSync(UI_CONFIG_PATH)
      ? JSON.parse(fs.readFileSync(UI_CONFIG_PATH, 'utf8'))
      : {};
    res.json(buildDataPaths(uiConfig));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── Data preview API ─────────────────────────────────────────────────────────
// Spawns gather-data.js --preview and returns the collected data as JSON.
// No extensions run — purely data collection for the UI Data tab.

app.get('/api/data/preview', (req, res) => {
  const envVars = readEnvFile();
  const child   = spawn('node', [path.join(projectRoot, 'gather-data.js'), '--preview'], {
    cwd: projectRoot,
    env: { ...process.env, ...envVars },
  });

  let out = '', err = '';
  child.stdout.on('data', d => out += d);
  child.stderr.on('data', d => err += d);

  child.on('close', () => {
    const marker = '__PREVIEW__';
    const idx    = out.indexOf(marker);
    if (idx === -1) {
      const firstErr = err.split('\n').find(l => l.trim()) || out || 'No output';
      return res.status(500).json({ error: firstErr });
    }
    try {
      res.json(JSON.parse(out.slice(idx + marker.length)));
    } catch {
      res.status(500).json({ error: 'Failed to parse preview data' });
    }
  });
  child.on('error', e => res.status(500).json({ error: e.message }));
});

// ─── Query test API ───────────────────────────────────────────────────────────
// Runs a single query in a fresh child process and returns field names + sample rows.
// Body: { wiqlTemplate, jiraTemplate, scope, workstreamName? }

app.post('/api/query/test', (req, res) => {
  const { wiqlTemplate = '', jiraTemplate = '', scope = 'workstream', workstreamName } = req.body;
  const envVars    = readEnvFile();
  const uiCfg      = fs.existsSync(UI_CONFIG_PATH) ? JSON.parse(fs.readFileSync(UI_CONFIG_PATH, 'utf8')) : {};
  const workstreams = uiCfg.workstreams || [];
  // For global scope: no workstream context — template vars resolve from env only
  const ws = scope === 'global'
    ? {}
    : (workstreamName ? workstreams.find(w => w.name === workstreamName) : null)
        || workstreams[0]
        || {};
  const provider    = (envVars.TEST_PROVIDER || process.env.TEST_PROVIDER || 'ado').toLowerCase();

  // Build a self-contained ES module script and pipe it via stdin
  const script = `
import { runQuery } from './providers/${provider}.js';
import { ADO_FIELD_MAP } from './config.js';
const ws       = ${JSON.stringify(ws)};
const scope    = ${JSON.stringify(scope)};
const provider = ${JSON.stringify(provider)};
let config;
if (provider === 'ado') {
  const wiql = ${JSON.stringify(wiqlTemplate)}.replace(/\\{\\{(\\w+)\\}\\}/g, (_, k) => ws[k] || process.env[k] || '');
  config = { wiql, fieldMap: ADO_FIELD_MAP };
} else {
  config = ${JSON.stringify(jiraTemplate)}.replace(/\\{\\{(\\w+)\\}\\}/g, (_, k) => ws[k] || process.env[k] || '');
}
const target = scope === 'global' ? { name: 'global', areaPath: '' } : ws;
try {
  const results = await runQuery(target, config);
  const out = {
    count:  results.length,
    fields: results[0] ? Object.keys(results[0]) : [],
    sample: results.slice(0, 3),
  };
  if (scope !== 'global') out.workstream = ws.name || '(none)';
  process.stdout.write(JSON.stringify(out));
} catch (e) {
  process.stdout.write(JSON.stringify({ error: e.message.split('\\n')[0] }));
}
`;

  const child = spawn('node', ['--input-type=module'], {
    cwd: projectRoot,
    env: { ...process.env, ...envVars },
  });

  let out = '', err = '';
  child.stdout.on('data', d => out += d);
  child.stderr.on('data', d => err += d);
  child.stdin.write(script);
  child.stdin.end();

  child.on('close', () => {
    try {
      res.json(JSON.parse(out));
    } catch {
      res.status(500).json({ error: err.split('\n').find(l => l.trim()) || out || 'No output' });
    }
  });
  child.on('error', err => res.status(500).json({ error: err.message }));
});

// ─── Start ────────────────────────────────────────────────────────────────────

app.listen(PORT, () => {
  console.log(`MMO Report UI running at http://localhost:${PORT}`);
  // Apply any saved schedule from ui-config.json on startup
  if (fs.existsSync(UI_CONFIG_PATH)) {
    try {
      const saved = JSON.parse(fs.readFileSync(UI_CONFIG_PATH, 'utf8'));
      applySchedule(saved.settings);
    } catch { /* ignore */ }
  }
});
