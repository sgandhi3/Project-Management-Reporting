# Setup Guide

This repo turns a manually-maintained ADO test execution / defect status report
(a PowerPoint deck) into one that regenerates itself from live Azure DevOps
data — optionally on a recurring schedule, with an AI-written narrative
summary and a pre-flight check that the numbers on every slide actually add
up before anything gets delivered.

It is **not a hosted app or service**. There is no server, no shared
infrastructure, and no central "instance" other people log into. Each team
that wants this clones the repo, brings their own ADO credentials, and runs
the setup below with an AI coding agent (this was built and is documented
against Claude Code) to configure it for their own project. The optional
recurring-automation phase runs as a scheduled task on *your own machine*
under *your own* Claude Code installation — it is not something a central
team operates on your behalf.

## What you need before starting

- **An ADO Personal Access Token** with Read access for Work Items + Test
  Management (Basic + Test Plans license), plus your ADO org and project name
- **Node 18+** (this repo is a Node/JavaScript project)
- **Python 3** — optional; the template-tokenization step (Phase 1, Step 3)
  can be written in Python or Node, your agent's choice
- **Your team's current report**, as an actual `.pptx` file — this becomes
  the template. Not a mockup or a different format; the real file people have
  been filling in by hand, with real (or realistic placeholder) numbers in it
- **If you plan to do Phase 2** (the recurring schedule): clone the repo to a
  permanent path, not a temp/scratch directory. The scheduled task hardcodes
  an absolute working directory that must still exist whenever it fires,
  in a completely separate future session.

## Quick start

1. Clone this repo, `npm install`
2. Copy `.env.example` → `.env` and fill in `ADO_ORG`, `ADO_PROJECT`,
   `ADO_PAT` — or just paste them in chat alongside the prompt below and let
   the agent write them in (it'll ask if anything's ambiguous)
3. Put your team's real report at the repo root as `source-report.pptx`
4. Paste the **Phase 1 prompt** below into Claude Code and let it work
   through the steps — it reads your report, discovers your ADO structure,
   asks you clarifying questions where the two don't cleanly match, and
   builds the token-substitution pipeline
5. Once Phase 1 is generating a correct report on demand, optionally paste
   the **Phase 2 prompt** to add the AI-written narrative and the recurring
   schedule

Each phase is one prompt. Paste it whole; don't split it up.

## Phase 1 — Automated Report Setup

```
## Automated Report Setup

You are configuring an automated PowerPoint report generation system
for a new project. This repo pulls live data from Azure DevOps and
generates a fully populated PPTX report. Your job is to configure it
for THIS project by reading the provided report and ADO structure.

### What's already in place
- `source-report.pptx` — the team's current manual report (repo root or wherever the user points you)
- `.env` — ADO credentials (populate before or during this prompt — see prerequisites)
- `temp2.pptx` — token-based template (you will replace this)
- `ui-config.json` — per-workstream suite IDs, plan ID, and area paths live here, not in `.env` (you will update the workstreams array)
- `variables.js` — token → data mappings (you will update this)
- `gather-data.js` — main entry point; reads `.env` + `ui-config.json` — you should not need to add new env vars

Note: there is no pre-existing `tokenize_pptx.py` or similar script in a fresh clone — you write it from scratch in Step 3. But do check scripts/ and extensions/ first: this repo has been reused across multiple project engagements, and often already has narrative-tokenization tooling (Phase 2, Steps 7–9) left over from a previous one. Adapt what's already there instead of creating parallel files with different names — see the note before Step 7.

### Step 1 — Read the existing report

source-report.pptx is a zip archive. Unzip it and read the XML
slides (ppt/slides/slide*.xml). Identify:

- How many slides and what each one covers
- Every workstream name as it appears in table row labels
- Every metric column (Total Cases, Executed, Passed, Failed,
  Blocked, Not Started, In Progress, Bugs, percentages)
- The executive summary slide structure (big stat boxes + table)
- Any special slides (defects by severity, by workstream, etc.)

Write a brief summary of what you found before continuing.

### Step 2 — Discover the ADO project structure

Read `.env` for credentials. Use the ADO REST API
(api-version=7.1) to:

1. List test plans to discover the plan ID — don't assume you already
   know it: GET /testplan/plans (a project usually has just one)
2. Confirm it and get the full suite hierarchy: GET /testplan/plans/{PLAN_ID}/suites?$expand=children
3. Get area paths: GET /wit/classificationnodes/areas?$depth=5

Build a mapping of each workstream found in the report to its:
- ADO suite ID (top-level suite under the plan root)
- Area path string (e.g. "ProjectName\WorkstreamName")
- Any sub-suite IDs for sub-workstreams that need separate tracking

Flag every mismatch between the report and the actual ADO tree, in
both directions — don't assume either side is right:
- A workstream in the report with no matching suite in ADO
- A top-level suite in ADO with no corresponding slide in the report
  (e.g. a workstream that hasn't started reporting yet)
- A workstream whose ADO sub-suites don't exactly match what's shown
  on its detail slide (extra sub-suites in ADO that aren't broken out
  in the report, or a sub-suite nested under an odd intermediate suite
  instead of directly under the workstream)

Ask the user how to handle each one before building the token mapping
— this materially changes what Step 4's totals should sum.

### Step 3 — Create the tokenized template

Write a script (Python + zipfile + regex, or Node + a zip library —
either works; regex-based text substitution on the raw slide XML is
safer than a full XML-tree round-trip, since re-serializing the whole
tree risks subtly altering namespaces or attribute order) that creates
a new `temp2.pptx` from `source-report.pptx`. There's no
pre-existing example of this specific script in a fresh clone — you're
writing it from scratch (but check scripts/ first per the note above,
in case a prior engagement left one behind).

Token naming convention:
- Per-workstream: {PREFIX}{METRIC}
  - PREFIX = 4–6 uppercase letters derived from workstream name
    (e.g. Enrollment → ENRL, EDI → EDI, Integration → INTG)
  - METRIC = TTC / ETC / EP / PTC / PP / FTC / FP / IPTC / NSTC / B
- Grand Total row: TTC / ETC / EP / PTC / PP / FTC / FP / IPTC / NSTC / TB
- Big stat boxes on exec slide: same tokens as Grand Total row

Rules:
- Replace ONLY numeric cells with {{TOKEN}} — preserve all text,
  colors, fonts, borders, and formatting exactly
- Percentage cells get EP/PP/FP tokens — but the token's value is just
  the bare number (e.g. 63). The replacement code does not
  append a % for you: keep the literal % character in the cell,
  right after the token (e.g. {{PDMEP}}%), same as it already
  appears in the source text
- Check each target cell's actual run structure before editing it —
  most numeric/stat-box cells already contain exactly one
  `<a:r><a:t>...</a:t></a:r>` run, so just replacing that
  existing `<a:t>` content in place is simplest and safest
- Only if a cell is genuinely empty (no `<a:r>` at all, just an
  `<a:endParaRPr/>`) do you need to insert a new `<a:r>` run —
  and it must go BEFORE any existing `<a:endParaRPr>` element
  in that paragraph

Verify by unzipping the output and confirming all {{TOKEN}}
placeholders are present and no endParaRPr ordering issues exist.

### Step 4 — Update variables.js

Edit `variables.js` to reflect this project (see the sub-suite
filtering rule above before wiring up totals):

1. Add a stat(d, ws, metric) call pattern for each workstream,
   where the workstream name exactly matches the ADO suite name
2. Map every token from Step 3 to the correct data accessor:
   - TTC: d => stat(d, 'WorkstreamName', 'planned')
   - ETC: d => stat(d, 'WorkstreamName', 'executed')
   - EP:  d => pct(stat(...executed), stat(...planned))
   - and so on for PTC/PP/FTC/FP/IPTC/NSTC/B
   - Critical: if Step 2 found a workstream whose ADO sub-suites
     don't all appear on its detail slide, do not use the
     provider's full recursive per-workstream stat (e.g. d.stats.Enrollment)
     for that workstream's totals — it silently includes the
     unreported sub-suites. Instead sum only the specific named
     sub-suite paths that are actually shown, with a
     sumStat(d, ws, [...paths], field)-style helper.
   - Per-sub-suite defect/bug counts are usually unavailable: ADO
     typically only has an area path at the top-level workstream, not
     per sub-suite. Hardcode those specific tokens to 0 rather than
     guessing — check if this codebase already has this precedent
     (a sub-suite token hardcoded to 0) before reinventing it.
3. Update Grand Total tokens to sum across all workstreams' totals —
   using each workstream's filtered total from the rule above, not a
   blanket sum of every raw d.stats entry
4. If any workstream gates its sub-plans on execution count (like
   Benefits in the original), add the activePlan filter helpers —
   use depth-2 key filtering: key.split(' / ').length === 2

### Step 5 — Update configuration

1. Update `ui-config.json`'s workstreams array with the suite
   ID, plan ID, and area path you discovered for each workstream — this
   is where per-workstream config lives in this codebase, not
   `.env.example`. `.env` only ever needs
   ADO_ORG / ADO_PROJECT / ADO_PAT — you should not
   need to invent new env vars for suite IDs or area paths.
2. Critical — check what the ppt output extension actually
   defaults PPTX_TEMPLATE to (commonly `temp.pptx`), and compare
   it against your real template's filename from Step 3
   (`temp2.pptx`). If they don't match, either rename the template
   to the default or set PPTX_TEMPLATE=./temp2.pptx explicitly in
   `.env`. If you skip this, report generation silently no-ops
   with "Template not found... skipping" instead of erroring — and
   Step 6's verification will falsely pass, since a file that was
   never created also has no unreplaced tokens to find.

### Step 6 — Run and verify

Run:
  node gather-data.js --out test-report.pptx

First confirm it actually printed something like "Report saved →
test-report.pptx" and the file exists on disk. If it printed
"Template not found... skipping" instead, the file was never created
— go fix the template path (Step 5) before treating a clean token
scan below as success, since an empty/missing file trivially has no
unreplaced tokens either.

Then inspect the output with Python:

  python3 -c "
  import zipfile, xml.etree.ElementTree as ET
  A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  with zipfile.ZipFile('test-report.pptx') as z:
      for name in z.namelist():
          if name.startswith('ppt/slides/slide'):
              txt = ''.join(t.text or '' for t in
                  ET.fromstring(z.read(name)).iter(f'{{{A}}}t'))
              if '{{' in txt:
                  print(name, '— UNREPLACED TOKENS FOUND')
  print('Scan complete')
  "

Report:
- Which slides look correct
- Any tokens still showing as {{TOKEN}} (unreplaced)
- Any slides with blank cells where data was expected
- The Grand Total and per-workstream numbers for a sanity check

Critical — math it out before declaring done. A token scan only
catches unreplaced placeholders; it does not catch numbers that are
individually wrong but still numbers. Add (or verify the provider
already has) an automatic check that every planned count partitions
exactly: passed + failed + blocked + paused + inProgress + notStarted === planned,
for every workstream and every sub-suite. Run it now and confirm it
passes with no mismatches. This exact check has already caught a real
bug in this codebase: different ADO orgs use different outcome
strings for "hasn't run yet" (one instance uses 'notExecuted',
another uses 'Active') — a provider that doesn't recognize a
given org's specific string can silently miscount it as executed
instead of not-started. Don't hardcode a specific "not started" string
whitelist; instead whitelist only the outcomes that unambiguously mean
the point was run (passed/failed/blocked/paused/inProgress) and
default everything else to notStarted — that's safe regardless of
which spelling this particular ADO instance uses.

Also cross-check that each workstream's total equals the sum of only
its shown sub-suites (per the Step 4 filtering rule), and that the
Grand Total equals the sum of the three workstream totals. If any of
this doesn't add up, the bug is upstream of the template — fix the
data pipeline, don't patch the symptom in variables.js.

Fix any issues found, then confirm the setup is complete.
```

## Phase 2 — Recurring Automation & Narrative Refresh (optional)

Only do this after Phase 1 is confirmed working.

```
## Phase 2 — Recurring Automation & Narrative Refresh

Phase 1 already tokenizes the numeric data and regenerates the report
on demand. Now make the free-text narrative sentences refresh from
live data too, and put the whole pipeline on a schedule.

### Note before Step 7 — check for leftover tooling first

This repo has been reused across multiple project engagements. Before
writing any new file in Steps 7–9, check scripts/ and
extensions/ for what a previous engagement may have already left
behind — it's common to find extensions/ai-narrative.js,
scripts/apply-ai-narrative-tokens.js, and a working
--narrative-data flag already fully built and already wired into
gather-data.js. If so, adapt the existing files' target
lists/keys to this project's slides and tokens — don't create parallel
files with different names (e.g. a new extensions/narrative.js
alongside an existing extensions/ai-narrative.js just leaves one
of them dead code).

### Step 7 — Find and tokenize stale narrative sentences

The source report almost certainly has hand-written status sentences
mixed in with the tables — an executive summary paragraph, a
per-workstream "Overall X testing is On Track/At Risk..." line,
defect detail callouts. These go stale the same way the numbers did.

Unzip the tokenized template and find every paragraph (<a:p>)
containing full sentences rather than table data. For each one:
- Note its exact text and which slide it's on
- Design a token name for it (e.g. {{AI_OVERALL_STATUS}},
  {{AI_<WORKSTREAM>_STATUS}})
- Replace the paragraph's runs with a single run containing the token,
  keeping the first run's formatting — collapse any other runs in that
  paragraph so PowerPoint doesn't reintroduce split runs

Write this as a small reusable script — check for an existing
scripts/apply-ai-narrative-tokens.js first (see the note above)
and adapt its target list rather than writing a new one from scratch —
that matches paragraphs by a short leading substring, not full-text
equality — invisible characters like non-breaking spaces make exact
string matching unreliable — so it can be re-run if the template
changes later.

### Step 8 — Add a narrative-refresh extension

Check for an existing extensions/ai-narrative.js first (see the
note above — very likely already present from a prior engagement) and
adapt its key list to this project's tokens, rather than creating a
new extensions/narrative.js. Whichever file it ends up being, it
does not generate any text itself — it only reads an
agent-override file named ai-narrative-input.json in the
project root (match this exact filename if one is already
.gitignore'd and referenced in the README — don't invent a
differently-named override file) and copies its values onto the
matching {{AI_...}} tokens, key by key:

- If the override file exists, use each key's value verbatim.
- If it doesn't exist, or a key is missing from it, leave that
  token blank. There is no generated fallback text — a plain
  manual run with no agent involved should not put invented,
  mechanically-assembled sentences into the deck.

This keeps the sentence quality bar at "a real Claude session actually
read the data and wrote this" or "blank" — never a synthetic middle
ground. Whatever writes the override file (Step 10) is responsible for
the actual judgment: status labels, defect themes, and never including
a person's name or an external client/facility name.

Wire the extension to run automatically before the main output
extension in gather-data.js, the same way any AI-summary
extension already does.

### Step 9 — Add a full-detail data-dump flag

Check gather-data.js first — this flag, and the wiring that runs
the narrative extension before the main output extension, may already
exist from a prior engagement (see the note above). If so, just
confirm it dumps full detail rather than rebuilding it. Otherwise: a
preview/summary flag (if the repo has one) may truncate item arrays to
a count — not enough detail to describe defect themes. Add a
--narrative-data flag to gather-data.js that dumps stats,
subStats, consolidatedData, and the FULL open-item arrays (title,
severity, owner, etc.) as JSON, skipping the output extensions. This
is what the scheduled task will read before writing the override file.

Critical — also include a tokens field: every non-AI_* token
from variables.js's VARIABLE_MAP, resolved against the same fetched
data (i.e. run each token's getter function and collect the results).
This is what makes narrative sentences accurate — the raw stats/
subStats in this dump are the FULL provider data, but the report may
deliberately show a filtered subset (per Step 4's sub-suite rule). If
the narrative-writing step sums raw stats/subStats itself instead
of quoting tokens, its sentences can describe numbers that don't
match what's actually printed in the deck.

Also make sure the automatic arithmetic-partition check from Step 6
(passed + failed + blocked + paused + inProgress + notStarted === planned)
runs as part of the normal pipeline, not just the one-time Step 6
verification — the scheduled task in Step 10 depends on it to detect a
bad run before it writes a misleading narrative.

### Step 9.5 — Fix text-overflow risk, and validate math on every slide

Two more one-time additions before scheduling anything, both concerned
with content that varies in length or value at runtime:

Formatting: check every shape whose text comes from a token or an
AI-narrative sentence (stat-box numbers, per-workstream status
sentences) for its <a:bodyPr> autofit setting. You will likely find
either no autofit at all (text silently overflows the box) or
<a:spAutoFit/> ("resize shape to fit text" — this grows the box,
which overlaps whatever sits below it, since PowerPoint shapes don't
push each other out of the way). Replace either with
<a:normAutofit/> ("shrink text to fit") on every such shape — this
keeps the shape's size, position, font, and colors exactly as
designed (no style change) and only reduces font size when content
would otherwise overflow. Write this as a small reusable, idempotent
script (check scripts/ first per the earlier note) that operates on
temp2.pptx directly, since the fix then applies to every future
report generated from it.

Math validation: write a second script that reads every number
actually printed across every slide of a generated report (not the
raw fetched data — this needs to catch token-wiring bugs too) and
cross-checks: sub-suite rows sum to their table's Grand Total row,
each detail slide's totals match the executive summary slide, the
executive summary's Grand Total equals the sum of the workstream rows,
any stat boxes match the Grand Total row, and every displayed
percentage matches what its underlying counts recompute to (allow ±1
for independent rounding). This is a different, complementary check
from Step 6/9's data-layer arithmetic check — that one validates the
fetched data before rendering; this one validates the rendered output
itself, catching mistakes introduced anywhere between the two.

### Step 10 — Create the recurring scheduled task

Critical — use a scheduling mechanism that is stored on disk and
survives across sessions. If Claude Code has both an in-session cron
helper and a proper scheduled-tasks feature (stored under something
like ~/.claude/scheduled-tasks/), use the durable one — an
in-session-only scheduler disappears the moment the current session
ends, and may auto-expire after about a week even if kept alive, which
makes it useless for a genuinely recurring weekly job. Verify which
one you actually created before trusting it.

This also means the project needs to live at a stable, permanent path
— not a temp/scratch directory scoped to the current session — since
the task's prompt hardcodes an absolute working directory that must
still exist whenever it fires, days or weeks later, in a completely
separate session. If the project is currently somewhere ephemeral,
move it to a permanent location first and point the task there.

Ask Claude Code to create this scheduled task (cron expression +
prompt). The prompt must be fully self-contained — each firing starts
a fresh session with no memory of this conversation. It should:

1. cd into the project directory
2. Run node gather-data.js --narrative-data, read the JSON. Check the
   console output (printed before the __NARRATIVE_DATA__ line) for the
   arithmetic-partition result from Step 6/9 — if it shows a MISMATCH
   instead of passing clean, stop here, skip to step 5, and report the
   exact mismatch as the failure. Otherwise, (using its own judgment,
   not a script) write ai-narrative-input.json with grounded,
   name-free sentences per the rules in Step 8 — match the exact
   filename extensions/ai-narrative.js actually reads.
   Every number in every sentence must be quoted directly from the
   `tokens` field of that JSON — never independently re-summed from
   `stats`/`subStats`, and never estimated. Before finalizing, read each
   sentence back against `tokens` and confirm every figure matches
   exactly, and that any relationship you stated in prose (e.g. "X of Y
   executed") is arithmetically consistent with those same values —
   don't let the model's own composition introduce an error that isn't
   in the data itself. If something doesn't add up, write a shorter,
   more conservative sentence rather than one with a mismatch in it.
3. Run node gather-data.js (no flags) to regenerate the report —
   this picks up the override file automatically. Confirm the console
   again shows the arithmetic check passing and "Report saved →..." —
   if not, treat this as a failed run.
4. Run the Step 9.5 formatting-fix script against the freshly generated
   report file (idempotent — the template already has the fix, so this
   is a cheap safety net, not a real change most weeks)
5. Run the Step 9.5 math-validation script against the same file. If it
   reports any inconsistency, do not proceed to deliver — a report
   with numbers that don't add up shouldn't go out even as a draft.
   Skip straight to the final notification step and report the exact
   mismatches found.
6. Deliver the report (see Step 11) — only if step 5 passed clean
7. Send exactly one push notification summarizing what happened,
   including whether the arithmetic check, formatting fix, and math
   validation all passed — this is the only way you'll know an
   unattended run succeeded or failed, so don't skip it, and send it
   last regardless of where the run stopped

### Step 11 — Choose and verify a delivery method

Check scripts/ first — send-email-local.js and/or
create-outlook-draft.js may already exist from a prior engagement on
this exact tenant, possibly with a comment already confirming the
outcome below. Don't assume SMTP will work regardless: many corporate
O365 tenants (this one included, if it's the same Deloitte tenant)
disable SmtpClientAuthentication entirely — no password (including
an app password) will get through, and you'll only find out by trying:

  node --env-file=.env -e "import('nodemailer').then(({default:nodemailer})=>{const t=nodemailer.createTransport({host:process.env.SMTP_HOST,port:587,secure:false,auth:{user:process.env.OUTLOOK_USER,pass:process.env.OUTLOOK_PASS}});t.verify().then(()=>console.log('OK')).catch(e=>console.error('FAIL:',e.message))})"

  Call transporter.verify() on a real nodemailer transport before
  relying on it for delivery. A failure citing
  "SmtpClientAuthentication is disabled" is a tenant policy, not a
  fixable credential problem — stop trying to fix credentials and move
  to a fallback.

Working fallbacks, in order of effort:
- A transactional email API you already have credentials for
  (SendGrid, Resend, Microsoft Graph's sendMail, etc.)
- Draft-and-send-manually: automate your desktop mail client (e.g.
  AppleScript for Microsoft Outlook on macOS) to open a pre-filled
  draft with the report attached — it never sends anything itself,
  a human clicks Send

Whichever you pick, verify it for real before trusting the schedule —
actually run it once and confirm a draft/email genuinely appeared.

### Critical — approve tool permissions once, manually

The very first time a scheduled task runs, it needs approval for every
tool it uses (running shell commands, writing files, automating a mail
client, sending notifications). If its first-ever firing is unattended,
it has no one to answer those prompts and silently stalls —
indistinguishable from "not running," with no error. Trigger it
manually once ("Run now") so you're present to approve each prompt;
those approvals are then stored on the task and applied automatically
to every future firing.
```

## Known limitations — read before rolling this out further

- **The delivery step (Phase 2, Step 11) is macOS + Outlook desktop
  specific** by default (it uses AppleScript to open a pre-filled draft).
  If your team is on Windows, or uses a different mail client, you'll need
  a different delivery mechanism — the working fallbacks listed in Step 11
  (a transactional email API, or a Windows-specific automation approach) are
  the starting point, but this needs to be adapted and verified for your
  environment before you trust the schedule.
- **This has not been through your firm's internal accelerator / security
  review process.** It handles a live ADO PAT (stored in a local `.env` file,
  which is fine for a single person's own machine but is not an
  enterprise-appropriate secrets story on its own) and pulls project/defect
  data that may be client-sensitive. Check with whoever runs your internal
  accelerator catalog about what's required before publishing this more
  broadly — this document does not constitute that review.
- **Every fresh setup configures its own instance.** There is no shared
  "generic" branch to fork from that's already been stripped of a specific
  project's workstream names, suite IDs, and template layout — Phase 1's job
  *is* that configuration step, done fresh against your own report and ADO
  project each time.
- **The narrative-writing step needs a live AI agent to run each time**
  (Phase 2, Step 10) — it is deliberately not a scripted template filler
  (see Phase 2, Step 8's rationale), so it only produces real sentences when
  an actual Claude Code session runs the scheduled task. There is no
  fallback "auto-generated" text; a run with no agent involved leaves those
  sentences blank rather than inventing filler.
