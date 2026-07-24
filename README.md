# Test Report Generator

Pulls test execution data and open bugs from your test management tool and populates a PowerPoint template automatically. Run it once and your weekly report is done.

Supports **Azure DevOps (ADO)**, **Jira**, and **Zephyr Scale** — switch between them with a single environment variable.

---

## Setup

### 1. Install dependencies

```bash
npm install
```

### 2. Choose your provider and configure your environment

Copy `.env.example` to `.env` and fill in the fields for the provider you're using. The key variable is:

```
TEST_PROVIDER=ado   # options: ado | jira | zephyr
```

#### Azure DevOps

```
ADO_ORG=your-org-name
ADO_PROJECT=your-project-name
ADO_PAT=your-personal-access-token

PLAN_ID=12345

SIT_SUITE_PDM=
SIT_SUITE_BENEFITS=
SIT_SUITE_ENROLLMENT=
SIT_SUITE_EDI=

PDMAreaPath=MyProject\PDM
BenefitsAreaPath=MyProject\Benefits
EnrollmentAreaPath=MyProject\Enrollment
EDIAreaPath=MyProject\EDI
```

Your PAT needs Read access for Test Plans and Work Items.

#### Jira

```
JIRA_DOMAIN=mycompany
JIRA_EMAIL=you@company.com
JIRA_API_TOKEN=your-jira-api-token

JIRA_PROJECT_PDM=PDM
JIRA_PROJECT_BENEFITS=BEN
JIRA_PROJECT_ENROLLMENT=ENR
JIRA_PROJECT_EDI=EDI

# Optional: leave blank to use the default JQL (issuetype = Test / issuetype = Bug AND statusCategory != Done)
JIRA_TEST_JQL_PDM=
JIRA_BUG_JQL_PDM=
```

Generate an API token at: https://id.atlassian.com/manage-profile/security/api-tokens

#### Zephyr Scale

```
ZEPHYR_TOKEN=your-zephyr-scale-api-token

ZEPHYR_CYCLE_PDM=PDM-C1
ZEPHYR_CYCLE_BENEFITS=BEN-C1
ZEPHYR_CYCLE_ENROLLMENT=ENR-C1
ZEPHYR_CYCLE_EDI=EDI-C1

# Zephyr fetches bugs from Jira, so you still need the Jira credentials
JIRA_DOMAIN=mycompany
JIRA_EMAIL=you@company.com
JIRA_API_TOKEN=your-jira-api-token

JIRA_PROJECT_PDM=PDM
# Optional per-workstream bug JQL:
JIRA_BUG_JQL_PDM=
```

### 3. Add your PowerPoint template

Place your slide deck in this folder and name it `temp.pptx`. If you prefer a different name or location, set `PPTX_TEMPLATE=/full/path/to/file.pptx` in your `.env`.

### 4. Run it

```bash
node generate-report.js
```

The filled-in PowerPoint is saved as `MMO_Report_YYYY-MM-DD.pptx` in this folder. You can change the output path with `--out`:

```bash
node generate-report.js --out ./reports/WeeklyReport.pptx
```

---

## Setting up your PowerPoint template

The script works like a mail merge — you put placeholder tokens in your slides and they get replaced with real values on each run. Here's how to set that up correctly.

**Steps:**

1. Open `temp.pptx` in PowerPoint.
2. Click into any text box where you want a value to appear.
3. Type the token exactly as shown — double curly braces, no spaces, case-sensitive. For example: `{{TTC}}`
4. **Important:** type it in one go without stopping, and don't change the font, color, or size partway through. PowerPoint internally splits text into separate chunks called "runs" whenever formatting changes, and if a token gets split across two runs it won't be replaced. The script has logic to fix most of these cases, but typing it fresh is the safest approach.
5. Save the file.

The script processes every slide, so you can use the same token on multiple slides and they'll all be populated.

---

## Available tokens

All tokens are defined in `variables.js`. Here's what's available by default.

### PDM

| Token | Meaning |
|---|---|
| `{{PDMTTC}}` | Planned test cases |
| `{{PDMETC}}` | Executed |
| `{{PDMNSTC}}` | Not started |
| `{{PDMIPTC}}` | In progress |
| `{{PDMPTC}}` | Passed |
| `{{PDMFTC}}` | Failed |
| `{{PDMTB}}` | Open bug count |

### Benefits

| Token | Meaning |
|---|---|
| `{{BTTC}}` | Planned test cases |
| `{{BETC}}` | Executed |
| `{{BNSTC}}` | Not started |
| `{{BIPTC}}` | In progress |
| `{{BPTC}}` | Passed |
| `{{BFTC}}` | Failed |
| `{{BTB}}` | Open bug count |

### Enrollment

| Token | Meaning |
|---|---|
| `{{ETTC}}` | Planned test cases |
| `{{EETC}}` | Executed |
| `{{ENSTC}}` | Not started |
| `{{EIPTC}}` | In progress |
| `{{EPTC}}` | Passed |
| `{{EFTC}}` | Failed |
| `{{ETB}}` | Open bug count |

### EDI

| Token | Meaning |
|---|---|
| `{{EDITTC}}` | Planned test cases |
| `{{EDIETC}}` | Executed |
| `{{EDINSTC}}` | Not started |
| `{{EDIIPTC}}` | In progress |
| `{{EDIPTC}}` | Passed |
| `{{EDIFTC}}` | Failed |
| `{{EDITB}}` | Open bug count |

### Overall (all workstreams combined)

| Token | Meaning |
|---|---|
| `{{TTC}}` | Total planned test cases |
| `{{ETC}}` | Total executed |
| `{{NSTC}}` | Total not started |
| `{{IPTC}}` | Total in progress |
| `{{PTC}}` | Total passed |
| `{{FTC}}` | Total failed |
| `{{TB}}` | Total open bugs |
| `{{PP}}` | Overall pass rate (whole number %) |
| `{{FP}}` | Overall fail rate (whole number %) |

### Other

| Token | Meaning |
|---|---|
| `{{Date}}` | Today's date in EST — MM/DD/YYYY |

---

## Adding a new workstream

**Step 1 — Add it to the WORKSTREAMS array in `generate-report.js`:**

Each workstream object holds config for all three providers. Only the fields that match your active `TEST_PROVIDER` are used — the rest are ignored.

```js
const WORKSTREAMS = [
  // existing entries...
  {
    name: 'Finance',    // this name is the key everywhere — must match d.stats.Finance in variables.js

    // ADO
    planId:     process.env.PLAN_ID,
    sitSuiteId: process.env.SIT_SUITE_FINANCE,
    areaPath:   process.env.FinanceAreaPath,

    // Jira
    projectKey: process.env.JIRA_PROJECT_FINANCE,
    testJql:    process.env.JIRA_TEST_JQL_FINANCE,
    bugJql:     process.env.JIRA_BUG_JQL_FINANCE,

    // Zephyr
    testCycleKey: process.env.ZEPHYR_CYCLE_FINANCE,
  },
];
```

**Step 2 — Add the relevant env vars to your `.env` for whichever provider you're using:**

```
# ADO
SIT_SUITE_FINANCE=88888
FinanceAreaPath=MyProject\Finance

# Jira
JIRA_PROJECT_FINANCE=FIN

# Zephyr
ZEPHYR_CYCLE_FINANCE=FIN-C1
```

**Step 3 — Add tokens for it to `variables.js`:**

```js
// Finance
FTTC:  d => d.stats.Finance.planned,
FETC:  d => d.stats.Finance.executed,
FNSTC: d => d.stats.Finance.notStarted,
FIPTC: d => d.stats.Finance.inProgress,
FPTC:  d => d.stats.Finance.passed,
FFTC:  d => d.stats.Finance.failed,
FTB:   d => d.bugs.Finance.length,
```

The workstream name in `d.stats.Finance` must match the `name` field in `WORKSTREAMS` exactly (case-sensitive).

**Step 4 — Add the tokens to your template** using the same process described in the PowerPoint setup section above.

---

## Adding or changing a variable

Open `variables.js`. Each line follows this pattern:

```js
TOKEN_NAME: d => <expression>,
```

The `d` object gives you access to everything the script fetched:

| What you write | What it gives you |
|---|---|
| `d.stats.PDM.planned` | PDM planned test case count |
| `d.stats.PDM.executed` | PDM executed count |
| `d.stats.PDM.passed` | PDM passed count |
| `d.stats.PDM.failed` | PDM failed count |
| `d.stats.PDM.notStarted` | PDM not started count |
| `d.stats.PDM.inProgress` | PDM in progress count |
| `d.bugs.PDM.length` | PDM open bug count |
| `d.consolidatedData.<field>` | Same fields but summed across all workstreams |
| `d.allBugs.total` | Total bugs across all workstreams |
| `d.allBugs.sev1` / `sev2` / `sev3` / `sev4` | Bugs broken down by severity |

Replace `PDM` with any workstream name from the `WORKSTREAMS` array.

**Examples:**

```js
// Pass rate for Benefits only
BEN_PP: d => d.stats.Benefits.executed
  ? Math.round((d.stats.Benefits.passed / d.stats.Benefits.executed) * 100)
  : 0,

// Critical bug count across all workstreams
CRIT_BUGS: d => d.allBugs.sev1,

// Static label — the function doesn't have to use d at all
CYCLE: () => 'SIT Cycle 3',
```

Once the token is in `variables.js`, add the matching `{{TOKEN_NAME}}` placeholder to your PowerPoint template and it will be filled in on the next run.

---

## Troubleshooting

**A token isn't being replaced in the output**
The most common cause is PowerPoint splitting the token internally. Delete the placeholder text and retype it from scratch in a single pass without changing any formatting mid-token. Avoid pasting from another source.

**All counts are 0 for a workstream**
Run with `DEBUG=true` to print the exact query being sent:
```bash
DEBUG=true node generate-report.js
```
For ADO: check that the area path in your `.env` matches what's in Project Settings → Area Paths exactly.
For Jira/Zephyr: paste the printed JQL into Jira's issue search to verify it returns results.

**`Unknown TEST_PROVIDER` error**
The value of `TEST_PROVIDER` in your `.env` must be exactly `ado`, `jira`, or `zephyr` (lowercase).

**401 error (ADO)**
Your PAT is expired or doesn't have the right scopes. Generate a new one with Read access for Work Items and Test Management.

**401 error (Jira / Zephyr)**
Check that `JIRA_EMAIL` and `JIRA_API_TOKEN` (or `ZEPHYR_TOKEN`) are set correctly. Jira API tokens are separate from your password — generate one at https://id.atlassian.com/manage-profile/security/api-tokens.

**404 error (ADO)**
`ADO_ORG` or `ADO_PROJECT` doesn't match exactly — both are case-sensitive.

**404 error (Zephyr)**
The `testCycleKey` value doesn't match an existing cycle. Check the Zephyr Scale UI for the exact key format (e.g. `PDM-C1`).

**Template file not found**
Make sure `temp.pptx` is in the same folder as `generate-report.js`, or set `PPTX_TEMPLATE` in your `.env` to the full path.
