# MMO Report Generator

Pulls test execution data and open bugs from your test management tool and automatically populates a PowerPoint and/or Excel report template. Run it once and your weekly report is done.

Supports **Azure DevOps (ADO)** and **Jira / Zephyr Scale** as data providers.

---

## How it works

```
.env (credentials)
ui-config.json (workstreams, settings, tokens)
     │
     ▼
gather-data.js          ← fetches stats + bugs from ADO or Jira
     │
     ├── extensions/ppt.js         ← fills {{tokens}} in your .pptx template
     ├── extensions/excel.js       ← fills {{tokens}} in your .xlsx template
     ├── extensions/ai-narrative.js ← fills {{AI_...}} narrative tokens from an agent-written override
     └── extensions/sharepoint.js  ← uploads output to SharePoint
```

Everything about workstreams, tokens, and output settings lives in `ui-config.json`. The only things that go in `.env` are credentials and secrets.

---

## Quick start

### 1. Install dependencies

```bash
npm install
```

### 2. Set your credentials in `.env`

Copy `.env.example` to `.env` and fill in only the credentials block for your provider:

```bash
cp .env.example .env
```

**Minimum required for ADO:**
```
TEST_PROVIDER=ado
ADO_ORG=your-org          # part after dev.azure.com/ in the URL
ADO_PROJECT=your-project
ADO_PAT=your-pat          # needs Read access for Work Items + Test Management
```

**Minimum required for Jira / Zephyr:**
```
TEST_PROVIDER=jira
JIRA_DOMAIN=mycompany     # part before .atlassian.net
JIRA_EMAIL=you@company.com
JIRA_API_TOKEN=your-token
ZEPHYR_TOKEN=your-zephyr-token   # only needed if using Zephyr for test stats
```

### 3. Configure workstreams in the UI

Start the config UI:

```bash
npm run ui
# → opens at http://localhost:3000
```

Go to the **Credentials** tab and configure your workstreams (suite IDs, area paths, project keys). This writes to `ui-config.json` — no need to put these values in `.env`.

### 4. Add your report templates

- **PowerPoint:** place your `.pptx` in this folder (default name: `temp.pptx`). Override with `PPTX_TEMPLATE=/path/to/file.pptx` in `.env`.
- **Excel:** place your `.xlsx` template here (default name: `template.xlsx`). Override with `EXCEL_TEMPLATE=/path/to/file.xlsx` in `.env`. Run `node create-excel-template.js` to generate a starter template with demo data.

### 5. Run the report

```bash
npm run generate
```

Output files are saved to the project folder:
- `MMO_Report_YYYY-MM-DD.pptx`
- `MMO_Report_YYYY-MM-DD.xlsx`
- `Summary_YYYY-MM-DD.txt` (AI summary, if enabled)

---

## The UI (`npm run ui`)

The browser-based UI runs a local server that lets you configure and run the report without touching any files directly.

| Tab | What it controls |
|---|---|
| **Data Preview** | Live view of fetched stats — bar charts, workstream table, sub-suite breakdown |
| **Settings** | Output formats (PPT / Excel / AI), schedule, fetch options |
| **Credentials** | ADO/Jira connection details, workstream suite IDs and area paths, template paths |
| **AI** | System prompt, user prompt suffix, context file upload, save-to-file toggle |

Changes made in the UI are saved to `ui-config.json` immediately. The next report run picks them up automatically.

### Scheduled runs

The UI supports automated scheduled runs via the Settings tab (cron expression). The scheduler only runs while the UI server is active (`npm run ui`).

---

## Setting up your PowerPoint template

The report works like a mail merge — you place `{{TOKEN}}` placeholders in your slides and they get replaced with live data on each run.

**Tips:**
- Type each token in one pass without stopping — PowerPoint splits text into internal "runs" when you pause or change formatting, which can break token detection.
- The same token can appear on multiple slides; all instances are replaced.
- Tokens are case-sensitive and must have no spaces inside the braces.

---

## Setting up your Excel template

Same `{{TOKEN}}` approach as PowerPoint. The extension reads your `.xlsx` template, replaces all tokens, then repopulates the data sheets (By Workstream, Sub-Suites, Bug Analysis) with live data.

To create a starter template with the correct sheet layout and column structure:

```bash
node create-excel-template.js
# creates template.xlsx in this folder
```

You can then open `template.xlsx`, add charts, adjust styling, and add any `{{TOKEN}}` cells you want. The data sheets are repopulated on each run; the Summary sheet tokens are replaced in-place.

---

## Available tokens

Tokens are defined in `ui-config.json` under `variableMappings`. Each entry has:
- `token` — the placeholder name used in templates (without the `{{}}`)
- `path` — a JavaScript expression evaluated against the data object `d`
- `description` — human-readable label shown in the UI

You can add, edit, or remove tokens from the UI (Settings → Variable Mappings) or directly in `ui-config.json`.

### Overall totals

| Token | Description |
|---|---|
| `{{TTC}}` | Total planned test cases (all workstreams) |
| `{{ETC}}` | Total executed |
| `{{PTC}}` | Total passed |
| `{{FTC}}` | Total failed |
| `{{NSTC}}` | Total not started |
| `{{IPTC}}` | Total in progress |
| `{{PP}}` | Overall pass % (passed / executed × 100) |
| `{{FP}}` | Overall fail % (failed / executed × 100) |
| `{{Date}}` | Report date — MM/DD/YYYY, Eastern time |

### PDM

| Token | Description |
|---|---|
| `{{PDMTTC}}` | Planned |
| `{{PDMETC}}` | Executed |
| `{{PDMPTC}}` | Passed |
| `{{PDMFTC}}` | Failed |
| `{{PDMNSTC}}` | Not started |
| `{{PDMIPTC}}` | In progress |
| `{{PDMIT2IT20BTC}}` | Iteration 2.0 — Blocked |
| `{{PDMIT2IT21BTC}}` | Iteration 2.1 — Blocked |
| `{{PDMIT2IT20PP}}` | Iteration 2.0 — Pass % |
| `{{PDMIT2IT20FP}}` | Iteration 2.0 — Fail % |
| `{{PDMIT2IT20EP}}` | Iteration 2.0 — Execution % |
| `{{PDMIT2IT21PP}}` | Iteration 2.1 — Pass % |
| `{{PDMIT2IT21FP}}` | Iteration 2.1 — Fail % |
| `{{PDMIT2IT21EP}}` | Iteration 2.1 — Execution % |
| `{{PDMIT2IT20CPBTC}}` | Iteration 2.0 / CPIMS — Blocked |
| `{{PDMIT2IT20ROBTC}}` | Iteration 2.0 / Rosters — Blocked |
| `{{PDMIT2IT20ANBTC}}` | Iteration 2.0 / Ancillary — Blocked |
| `{{PDMIT2IT21CPBTC}}` | Iteration 2.1 / CPIMS — Blocked |
| `{{PDMIT2IT21ROBTC}}` | Iteration 2.1 / Rosters — Blocked |
| `{{PDMIT2IT21ANBTC}}` | Iteration 2.1 / Ancillary — Blocked |

### Benefits

| Token | Description |
|---|---|
| `{{BTTC}}` | Planned |
| `{{BETC}}` | Executed |
| `{{BPTC}}` | Passed |
| `{{BFTC}}` | Failed |
| `{{BNSTC}}` | Not started |
| `{{BIPTC}}` | In progress |
| `{{BENEBTC}}` | Blocked |
| `{{BENEPP}}` | Pass % |
| `{{BENEFP}}` | Fail % |
| `{{BENEEP}}` | Execution % |
| `{{BENESIHMBTC}}` | Signature HMO — Blocked |
| `{{BENEACPPBTC}}` | Access PPO (Premium PPO INN) — Blocked |
| `{{BENEPRPPBTC}}` | Premium PPO — Blocked |

### Enrollment

| Token | Description |
|---|---|
| `{{ETTC}}` | Planned |
| `{{EETC}}` | Executed |
| `{{EPTC}}` | Passed |
| `{{EFTC}}` | Failed |
| `{{ENSTC}}` | Not started |
| `{{EIPTC}}` | In progress |
| `{{ENROBTC}}` | Blocked |
| `{{ENROPP}}` | Pass % |
| `{{ENROFP}}` | Fail % |
| `{{ENROEP}}` | Execution % |

### EDI

| Token | Description |
|---|---|
| `{{EDITTC}}` | Planned |
| `{{EDIETC}}` | Executed |
| `{{EDIPTC}}` | Passed |
| `{{EDIFTC}}` | Failed |
| `{{EDINSTC}}` | Not started |
| `{{EDIIPTC}}` | In progress |
| `{{EDIBTC}}` | Blocked |
| `{{EDIPP}}` | Pass % |
| `{{EDIFP}}` | Fail % |
| `{{EDIEP}}` | Execution % |

---

## Test execution statuses

| Status | Counted as Executed? | Notes |
|---|---|---|
| Passed | Yes | |
| Failed | Yes | |
| In Progress | Yes | |
| Blocked | **No** | Test was blocked from running — tracked separately |
| Not Started | No | |

`planned = executed + notStarted + inProgress + blocked`

---

## Adding a new workstream

**Option A — UI (recommended):**
Go to Credentials → Workstreams → Add Workstream. Fill in the name, suite ID, and area path. Save.

**Option B — Edit `ui-config.json` directly:**

```json
{
  "name": "Finance",
  "planId": "20140",
  "sitSuiteId": "99999",
  "areaPath": "HRP\\Testing\\Finance",
  "projectKey": "",
  "testCycleKey": ""
}
```

Then add tokens for it in `variableMappings`:

```json
{ "token": "FINTTC",  "path": "d.stats.Finance.planned",    "description": "Finance — Planned" },
{ "token": "FINETC",  "path": "d.stats.Finance.executed",   "description": "Finance — Executed" },
{ "token": "FINPTC",  "path": "d.stats.Finance.passed",     "description": "Finance — Passed" }
```

The workstream `name` is the key used everywhere — it must match exactly (case-sensitive) in `d.stats.<name>` expressions.

---

## The data object (`d`)

Token expressions have access to the full data object. Here's what's available:

| Expression | What it returns |
|---|---|
| `d.stats.PDM.planned` | PDM planned test case count |
| `d.stats.PDM.executed` | PDM executed count |
| `d.stats.PDM.passed` | PDM passed count |
| `d.stats.PDM.failed` | PDM failed count |
| `d.stats.PDM.notStarted` | PDM not started count |
| `d.stats.PDM.inProgress` | PDM in progress count |
| `d.stats.PDM.blocked` | PDM blocked count |
| `d.subStats.PDM['Iteration 2 / Iteration 2.0'].executed` | Sub-suite stat (breadcrumb path) |
| `d.consolidatedData.planned` | Total planned across all workstreams |
| `d.bugs.PDM` | Array of open PDM bugs |
| `d.bugs.PDM.length` | PDM open bug count |
| `d.bugsTotal` | Total open bugs across all workstreams |
| `d.bugsBySeverity['1 - Critical']` | Count of Critical bugs |
| `d.bugsByPriority['1']` | Count of P1 bugs |
| `d._aiNarratives.overallStatus` | AI-refreshed narrative sentence (see AI Narrative Refresh below) |

Replace `PDM` with any workstream name. Sub-suite paths use the exact breadcrumb string from ADO.

---

## Narrative Refresh

The `temp2.pptx` template has several free-text status sentences (executive summary bullets, each workstream's "Overall ... testing is On Track/At Risk ..." line, and the PDM defect-detail cells) that used to be hand-typed and went stale week to week. These are now `{{AI_...}}` tokens (see `extensions/ai-narrative.js` for the full list).

**These tokens are only ever filled in one way: an agent-written override.** If `ai-narrative-input.json` exists in the project root, its values are used verbatim, key by key. The scheduled task writes this itself: it runs `node gather-data.js --narrative-data` to get the full current data (including complete defect lists), reads it with its own judgment (it's already a live Claude Code session, so this needs no separate API key or billing), and writes real sentences — grounded in the data, no invented facts, and never naming an individual defect owner. See the task's prompt (`~/.claude/scheduled-tasks/mmo-weekly-report-friday/SKILL.md`) for the exact instructions it follows.

There is deliberately **no generated fallback text**. A plain `npm run generate` with no agent involved (or any key the agent didn't write) leaves that `{{AI_...}}` token blank in the deck. This runs automatically whenever `ppt` is in the output formats.

If you retire the template and rebuild `temp2.pptx` from scratch, re-run `node scripts/apply-ai-narrative-tokens.js` to re-tokenize the new file's hardcoded sentences (edit the `TARGETS` list in that script first to match the new wording).

---

## SharePoint upload

Add `sharepoint` to your output formats and set these in `.env`:

```
SHAREPOINT_TENANT_ID=       # Azure AD → Overview → Directory (tenant) ID
SHAREPOINT_CLIENT_ID=       # App registration → Application (client) ID
SHAREPOINT_CLIENT_SECRET=   # Client secret value
SHAREPOINT_SITE_URL=        # e.g. https://myorg.sharepoint.com/sites/Weekly-Reports
SHAREPOINT_FOLDER=          # e.g. Shared Documents/Weekly Status
```

**Azure AD setup (one-time):**
1. portal.azure.com → Azure Active Directory → App registrations → New registration
2. API permissions → Microsoft Graph → Application → `Sites.ReadWrite.All`
3. Grant admin consent
4. Certificates & secrets → New client secret → copy the Value

---

## Troubleshooting

**A token isn't being replaced in the PowerPoint**
Delete the placeholder and retype it from scratch in one pass without changing formatting mid-token. PowerPoint splits text internally, which can break token detection.

**All counts are 0 for a workstream**
Check the suite ID and area path in the UI (Credentials tab). For ADO, the area path must match exactly what's in Project Settings → Area Paths. Try running with `DEBUG=true` to print the raw queries.

**`Unknown OUTPUT_FORMAT` error**
The format name must match a file in `extensions/`. Available: `ppt`, `excel`, `sharepoint`.

**401 error (ADO)**
Your PAT is expired or missing scopes. Generate a new one at `https://dev.azure.com/{org}/_usersSettings/tokens` with Read access for Work Items and Test Management.

**401 error (Jira / Zephyr)**
Verify `JIRA_EMAIL` and `JIRA_API_TOKEN`. Jira API tokens are separate from your password — generate one at https://id.atlassian.com/manage-profile/security/api-tokens.

**404 error (ADO)**
`ADO_ORG` and `ADO_PROJECT` are case-sensitive and must match exactly.

**Excel merge error on run**
This can happen if your template has merged cells in a range the extension tries to repopulate. The extension handles this automatically — if you see it, check that `EXCEL_TEMPLATE` points to the correct file.

**Template file not found**
Default PPT template: `temp.pptx` in the project folder. Default Excel template: `template.xlsx`. Override with `PPTX_TEMPLATE` or `EXCEL_TEMPLATE` in `.env`.
