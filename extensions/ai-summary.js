// AI Status Summary extension
//
// Calls Claude to generate an executive status summary from:
//   1. The collected data object (test stats, bugs, workstream counts)
//   2. An optional external notes source — a file path (from ui-config or env var)
//      or a URL that returns plain text or CSV.
//
// Config (via .env):
//   ANTHROPIC_API_KEY  — your Anthropic API key
//   SUMMARY_NOTES_URL  — URL to your online notes/Excel export (optional but recommended)
//   SUMMARY_NOTES_FILE — path to a local notes/context file (optional)
//   SUMMARY_OUTPUT     — where to save the generated text (default: Summary_YYYY-MM-DD.txt)
//
// Config (via ui-config.json settings.aiSettings):
//   systemPrompt       — overrides the built-in system prompt
//   userPromptSuffix   — appended to the data/notes block
//   contextFile        — path to a context file (overrides SUMMARY_NOTES_FILE)
//   saveToFile         — if false, skip writing to disk (default: true)

import fs   from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import Anthropic from '@anthropic-ai/sdk';

const __dirname2   = path.dirname(fileURLToPath(import.meta.url));
const UI_CONFIG_PATH = path.join(__dirname2, '..', 'ui-config.json');

const outputPath = process.env.SUMMARY_OUTPUT
  || path.join(process.cwd(), `Summary_${new Date().toISOString().slice(0, 10)}.txt`);

// ─── Read ai settings from ui-config ─────────────────────────────────────────

function readAiSettings() {
  try {
    if (fs.existsSync(UI_CONFIG_PATH)) {
      const cfg = JSON.parse(fs.readFileSync(UI_CONFIG_PATH, 'utf8'));
      return cfg.settings?.aiSettings || {};
    }
  } catch { /* ignore */ }
  return {};
}

// ─── Notes loader ─────────────────────────────────────────────────────────────

async function fetchNotes(aiSettings) {
  // ui-config contextFile takes precedence over SUMMARY_NOTES_FILE env var
  const configFile = aiSettings.contextFile;
  const envFile    = process.env.SUMMARY_NOTES_FILE;
  const url        = process.env.SUMMARY_NOTES_URL;

  const file = configFile || envFile;

  if (file) {
    const resolved = path.resolve(file);
    if (!fs.existsSync(resolved)) {
      console.warn(`  ⚠  Context file not found: ${resolved}`);
      // Fall through to URL if available
    } else {
      console.log(`  Notes loaded from file: ${resolved}`);
      return fs.readFileSync(resolved, 'utf8');
    }
  }

  if (url) {
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      console.log(`  Notes fetched from URL: ${url}`);
      return await res.text();
    } catch (e) {
      console.warn(`  ⚠  Could not fetch notes from SUMMARY_NOTES_URL — ${e.message}`);
      return null;
    }
  }

  return null;
}

// ─── Data formatter ───────────────────────────────────────────────────────────

function formatDataForPrompt(data) {
  const lines = [];

  // Per-workstream execution table
  lines.push('TEST EXECUTION STATUS');
  lines.push('─'.repeat(60));
  lines.push('Workstream | Planned | Executed | Passed | Failed | Not Started | In Progress | Blocked');
  for (const [ws, s] of Object.entries(data.stats)) {
    const pct = s.executed ? Math.round((s.passed / s.executed) * 100) : 0;
    lines.push(`${ws} | ${s.planned} | ${s.executed} | ${s.passed} | ${s.failed} | ${s.notStarted} | ${s.inProgress} | ${s.blocked ?? 0} (${pct}% pass rate)`);
  }
  const c = data.consolidatedData;
  const totalPct = c.executed ? Math.round((c.passed / c.executed) * 100) : 0;
  lines.push(`TOTAL | ${c.planned} | ${c.executed} | ${c.passed} | ${c.failed} | ${c.notStarted} | ${c.inProgress} | ${c.blocked ?? 0} (${totalPct}% pass rate)`);
  lines.push('');

  // Bug summary — finds any top-level key containing "bug" that's an object (groupBy results)
  const bugKeys = Object.keys(data).filter(k =>
    k.toLowerCase().includes('bug') && typeof data[k] === 'object' && data[k] !== null && !Array.isArray(data[k])
  );
  if (bugKeys.length) {
    lines.push('BUG SUMMARY');
    lines.push('─'.repeat(60));
    for (const key of bugKeys) {
      const val = data[key];
      if (val.total !== undefined) lines.push(`${key}.total: ${val.total}`);
      if (val.severity) {
        const { total, sev1, sev2, sev3, sev4 } = val.severity;
        lines.push(`  By severity — Critical: ${sev1}, High: ${sev2}, Medium: ${sev3}, Low: ${sev4} (total: ${total})`);
      }
      if (val.priority) {
        const { pri1, pri2, pri3, pri4 } = val.priority;
        lines.push(`  By priority — P1: ${pri1}, P2: ${pri2}, P3: ${pri3}, P4: ${pri4}`);
      }
    }

    // Per-workstream bug counts
    if (data.bugs) {
      lines.push('  Per workstream open bugs:');
      for (const [ws, bugs] of Object.entries(data.bugs)) {
        lines.push(`    ${ws}: ${Array.isArray(bugs) ? bugs.length : '?'}`);
      }
    }
    if (data.closedBugs) {
      lines.push('  Per workstream closed bugs:');
      for (const [ws, bugs] of Object.entries(data.closedBugs)) {
        lines.push(`    ${ws}: ${Array.isArray(bugs) ? bugs.length : '?'}`);
      }
    }
    lines.push('');
  }

  return lines.join('\n');
}

// ─── Prompt builder ───────────────────────────────────────────────────────────

const BUILT_IN_SYSTEM_PROMPT = `You are a senior QA manager writing a weekly status update for a SIT (System Integration Testing) program. Be concise, factual, and professional.`;

const BUILT_IN_USER_TEMPLATE = (dataBlock, notesBlock, suffix) =>
  `Below is the current test execution data and bug summary, followed by any additional project notes.

${dataBlock}${notesBlock}
---

Write a concise status update in exactly two sections:

**Executive Summary:**
2–4 sentences covering overall SIT progress, what's on track, and any notable workstream-specific highlights or blockers. Reference specific workstreams by name. Keep it factual and professional.

**Bug Triage:**
1–2 sentences summarizing the open bug situation: total count, severity/priority breakdown, and triage status. Reference any owners or next steps if the notes mention them.

Do not add headers beyond the two above. Do not invent facts not supported by the data. Match the tone: direct, professional, past-tense for completed work, present/future for ongoing.${suffix ? '\n\n' + suffix : ''}`;

function buildMessages(data, notes, aiSettings) {
  const dataBlock  = formatDataForPrompt(data);
  const notesBlock = notes
    ? `\nADDITIONAL NOTES / PROJECT CONTEXT\n${'─'.repeat(60)}\n${notes.slice(0, 8000)}\n`
    : '';
  const suffix = aiSettings.userPromptSuffix || '';

  const userContent = BUILT_IN_USER_TEMPLATE(dataBlock, notesBlock, suffix);
  return { userContent };
}

// ─── Provider interface ───────────────────────────────────────────────────────

export async function generate(data) {
  if (!process.env.ANTHROPIC_API_KEY) {
    console.error('\n❌  ANTHROPIC_API_KEY not set — skipping AI summary.');
    return;
  }

  console.log('\nGenerating AI status summary...');

  const aiSettings = readAiSettings();
  const notes      = await fetchNotes(aiSettings);
  const { userContent } = buildMessages(data, notes, aiSettings);
  const client     = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

  // Use custom system prompt if configured, otherwise use built-in
  const systemPrompt = aiSettings.systemPrompt || BUILT_IN_SYSTEM_PROMPT;

  let summary;
  try {
    const response = await client.messages.create({
      model:      'claude-sonnet-4-5',
      max_tokens: 512,
      system:     systemPrompt,
      messages:   [{ role: 'user', content: userContent }],
    });
    summary = response.content[0].text;
  } catch (e) {
    console.error(`\n❌  Claude API error — ${e.message.split('\n')[0]}`);
    return;
  }

  // Store in data for downstream extensions (ppt, excel)
  data._aiSummary = summary;

  const divider = '─'.repeat(64);
  console.log(`\n${divider}`);
  console.log(' AI STATUS SUMMARY');
  console.log(divider);
  console.log('__AI_START__');
  console.log('\n' + summary + '\n');
  console.log('__AI_END__');
  console.log(divider + '\n');

  // Write to file unless explicitly disabled
  if (aiSettings.saveToFile !== false) {
    fs.writeFileSync(outputPath, summary);
    console.log(`Summary saved → ${outputPath}`);
  }
}
