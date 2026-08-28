// AI Narrative Refresh extension
//
// The .pptx template has several free-text status sentences (executive summary
// bullets, per-workstream "Overall ... testing is On Track/At Risk ..." lines,
// and PDM defect-detail cells) that go stale week to week — hardcoded dates,
// counts, and plan names. This extension asks Claude to rewrite each one from
// this week's actual data and stores the results on `data._aiNarratives`,
// which variables.js exposes as {{AI_...}} tokens for extensions/ppt.js to
// substitute — the same pattern ai-summary.js uses for data._aiSummary.
//
// Config (via .env): ANTHROPIC_API_KEY, SUMMARY_NOTES_FILE / SUMMARY_NOTES_URL
// (same optional context source ai-summary.js uses, for any qualitative detail
// — e.g. named root causes — that isn't present in the raw ADO fields).
//
// If ANTHROPIC_API_KEY is missing, or the API call fails, every {{AI_...}}
// token falls back to an empty string rather than blocking report generation.

import Anthropic from '@anthropic-ai/sdk';
import { readAiSettings, fetchNotes } from './_ai-shared.js';

const NARRATIVE_KEYS = [
  'overallStatus',
  'pdmIterationUpdate',
  'pdmDefectSummary',
  'benefitsUpdate',
  'pdmCursoryStatus',
  'benefitsStatus',
  'enrollmentStatus',
  'ediStatus',
  'pdmDataQualityDetail1',
  'pdmDataQualityDetail2',
  'pdmDataMappingDetail',
];

function pct(n, d) { return d ? Math.round((n / d) * 100) : 0; }

function workstreamLine(name, s) {
  if (!s) return `${name}: no data`;
  return `${name} — planned ${s.planned}, executed ${s.executed} (${pct(s.executed, s.planned)}%), `
    + `passed ${s.passed} (${pct(s.passed, s.executed)}% of executed), failed ${s.failed}, `
    + `blocked ${s.blocked ?? 0}, in progress ${s.inProgress}, not started ${s.notStarted}`;
}

function bugLines(bugs) {
  if (!bugs?.length) return '  (none open)';
  return bugs.map(b => `  #${b.id} [${b.severity || 'Unknown'}/${b.priority || '—'}] "${b.title}" — owner: ${b.owner}, state: ${b.state}`).join('\n');
}

function buildDataBlock(data) {
  const lines = [];
  lines.push('TEST EXECUTION BY WORKSTREAM');
  for (const [ws, s] of Object.entries(data.stats)) lines.push(workstreamLine(ws, s));
  lines.push('');

  const pdmBugs        = data.bugs?.PDM ?? [];
  const benefitsBugs   = data.bugs?.Benefits ?? [];
  const enrollmentBugs = data.bugs?.Enrollment ?? [];
  const ediBugs        = data.bugs?.EDI ?? [];
  const pdmHigh   = pdmBugs.filter(b => b.severity === '2 - High' || b.severity === '1 - Critical');
  const pdmMedium = pdmBugs.filter(b => b.severity === '3 - Medium');

  lines.push('OPEN DEFECTS');
  lines.push(`PDM: ${pdmBugs.length} total (${pdmHigh.length} critical/high, ${pdmMedium.length} medium)`);
  lines.push(bugLines(pdmBugs));
  lines.push(`Benefits: ${benefitsBugs.length} total`);
  lines.push(bugLines(benefitsBugs));
  lines.push(`Enrollment: ${enrollmentBugs.length} total`);
  lines.push(bugLines(enrollmentBugs));
  lines.push(`EDI: ${ediBugs.length} total`);
  lines.push(bugLines(ediBugs));
  lines.push('');

  lines.push('PDM CRITICAL/HIGH SEVERITY DEFECTS (for the "Data Quality" detail cell)');
  lines.push(bugLines(pdmHigh));
  lines.push('');
  lines.push('PDM MEDIUM SEVERITY DEFECTS (for the "Data Mapping" detail cell)');
  lines.push(bugLines(pdmMedium));

  return lines.join('\n');
}

const SYSTEM_PROMPT = `You are a senior QA program manager updating a weekly SIT (System Integration Testing) status deck. You rewrite specific status sentences using only the data given to you — you never invent dates, counts, names, or root causes that aren't in the data or notes provided. When a category has zero relevant items, say so plainly instead of fabricating detail.`;

function buildUserPrompt(dataBlock, notesBlock, todayLabel) {
  return `Today's date for the report: ${todayLabel}

${dataBlock}
${notesBlock}
---

Rewrite the following status sentences/cells for this week's deck, using ONLY the data above (and the notes, if any, for qualitative context like named root causes or owners). Each one replaces a specific hardcoded sentence in the deck that is now out of date. Match the length and tone of a typical entry (1-3 sentences, direct, professional, present the facts plainly — do not editorialize beyond what "on track / at risk / off track / behind plan" the data supports).

Guidance for the on track / at risk / off track / behind plan judgment: "on track" = execution proceeding with no major blockers and pass rate healthy; "at risk" = meaningful defects or slippage but recoverable; "off track" / "behind plan" = execution stalled, blocked, or materially behind. Base the call on the executed/passed/failed/blocked numbers given.

Return STRICT JSON only, with exactly these keys (all string values, no markdown, no extra keys, no trailing commentary):

{
  "overallStatus": "1-2 sentence overall SIT program status across all workstreams",
  "pdmIterationUpdate": "1 sentence on current PDM iteration progress/timeline",
  "pdmDefectSummary": "1-2 sentences on the PDM defect situation (types of defects, how many resolved vs remaining, any target date implied by the data)",
  "benefitsUpdate": "1-2 sentences on Priority Benefits testing progress (pass rate, what's next)",
  "pdmCursoryStatus": "1-2 sentences: overall PDM Cursory status line, defect count, and who owns them (HE vs MMO) based on defect data/notes",
  "benefitsStatus": "1-2 sentences: overall Benefits SIT status line and current open defect count/theme",
  "enrollmentStatus": "1-2 sentences: overall Enrollment SIT status line, defect count, and defect themes",
  "ediStatus": "1-2 sentences: overall EDI SIT status line, defect count, and defect themes",
  "pdmDataQualityDetail1": "1 sentence describing the nature of the current PDM critical/high-severity defects (first half of a two-line explanation)",
  "pdmDataQualityDetail2": "1 sentence continuing that explanation with any second distinct theme, or empty string if there is only one theme",
  "pdmDataMappingDetail": "1-2 sentences describing the nature of the current PDM medium-severity defects"
}`;
}

export async function generate(data) {
  data._aiNarratives = data._aiNarratives || {};

  if (!process.env.ANTHROPIC_API_KEY) {
    console.error('\n❌  ANTHROPIC_API_KEY not set — skipping AI narrative refresh (AI_* tokens will be blank).');
    return;
  }

  console.log('\nRefreshing AI-generated status sentences...');

  const aiSettings = readAiSettings();
  const notes      = await fetchNotes(aiSettings);
  const notesBlock = notes ? `\nADDITIONAL NOTES / PROJECT CONTEXT\n${'─'.repeat(60)}\n${notes.slice(0, 8000)}\n` : '';
  const dataBlock  = buildDataBlock(data);
  const todayLabel = new Date().toLocaleDateString('en-US', {
    month: 'long', day: 'numeric', year: 'numeric', timeZone: 'America/New_York',
  });

  const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

  let parsed;
  try {
    const response = await client.messages.create({
      model:      'claude-sonnet-4-5',
      max_tokens: 1024,
      system:     SYSTEM_PROMPT,
      messages:   [{ role: 'user', content: buildUserPrompt(dataBlock, notesBlock, todayLabel) }],
    });
    const text = response.content[0].text.trim();
    const jsonText = text.replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
    parsed = JSON.parse(jsonText);
  } catch (e) {
    console.error(`\n❌  AI narrative generation failed — ${e.message.split('\n')[0]} (AI_* tokens will be blank).`);
    return;
  }

  for (const key of NARRATIVE_KEYS) {
    data._aiNarratives[key] = typeof parsed[key] === 'string' ? parsed[key].trim() : '';
  }

  console.log('AI narrative sentences:');
  for (const key of NARRATIVE_KEYS) {
    console.log(`  ${key}: ${data._aiNarratives[key] ? data._aiNarratives[key].slice(0, 90) + (data._aiNarratives[key].length > 90 ? '…' : '') : '(empty)'}`);
  }
}
