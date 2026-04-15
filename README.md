# AI Job Application System

Automated Seek / LinkedIn job pipeline for:
- scraping jobs and JDs
- reusing logged-in browser sessions
- tailoring resumes and cover letters
- scoring ATS match quality
- tracking results in Excel / CSV / Markdown
- assisting with manual application flow

This codebase keeps the execution layer practical and stable, while the generation layer is now decision-driven:

1. Parse JD
2. Classify role family
3. Analyze base resume
4. Select high-signal profile facts
5. Inject persona-specific evidence
6. Generate role-locked resume
7. Score resume with weighted ATS logic
8. Regenerate once if needed
9. Generate a separate cover letter

## Project Layout

### Core pipeline modules

- `scripts/jd_parser.py`
  - Cleans JD text
  - Removes branding / benefits / noise
  - Extracts keywords, tools, qualifications, responsibilities, and domain terms

- `scripts/role_classifier.py`
  - Classifies target role family
  - Current important families:
    - `accounting_reporting`
    - `accounting_systems`
    - `business_analyst`
    - `data_analyst`
    - `business_graduate`
    - `sales_bd`
    - `operations`
    - `consulting`

- `scripts/resume_analyzer.py`
  - Parses the base resume into sections
  - Scores each experience block
  - Marks roles as `KEEP`, `BOOST`, or `DOWNPLAY`
  - Preserves all role headers so generation never deletes real experience

- `scripts/profile_filter.py`
  - Ranks profile facts and external facts
  - Keeps only compact, role-relevant prompt evidence
  - Downranks noisy or irrelevant facts

- `scripts/finance_signal_injector.py`
  - Detects finance evidence and software evidence
  - Injects AP / AR / invoicing / reconciliations / month-end / reporting / systems signals
  - Resolves accounting sub-persona:
    - `accounting_reporting`
    - `accounting_systems`

- `scripts/resume_generator.py`
  - Persona-aware title generator
  - Persona-aware summary builder
  - Persona-aware skills builder
  - Persona-aware bullet prioritization and rewriting
  - Fallback ATS-safe resume renderer

- `scripts/ats_scorer.py`
  - Scores only job-relevant signals
  - Uses weighted role-aware logic
  - Rewards contextual usage in experience more than flat keyword lists
  - Returns missing tier-1 keywords, weak sections, and refinements

- `scripts/cover_letter_generator.py`
  - Generates a separate cover letter
  - Never leaks ATS internal notes into user-facing output

- `scripts/pipeline_orchestrator.py`
  - Main decision controller
  - Owns the generation order:
    - parse JD
    - classify role
    - analyze resume
    - filter profile facts
    - inject finance signals if relevant
    - generate resume
    - score
    - refine once if score is low
    - generate cover letter

### Main workflow script

- `scripts/job_workflow.py`
  - CLI entrypoint
  - Handles:
    - `auth`
    - `scrape`
    - `run`
    - `apply`
    - `tracker`
    - `hydrate-profile`
    - `scan-portals`

### Config and profile data

- `config/candidate_profile.json`
  - Highest-priority candidate facts
  - Includes structured finance capability evidence
  - Includes accounting software exposure evidence

- `config/external_profile_facts.json`
  - Supplemental verified facts
  - Hydrated from the Digital Twin project

- `config/settings.yaml`
  - Runtime settings

- `config/portals.yaml`
  - Greenhouse / Ashby / Lever company portal scan targets

### Inputs and outputs

- `jobs/jobs.csv`
  - Scraped or manually provided job list

- `outputs/applications/`
  - Generated resume and cover letter `.docx` files

- `outputs/reports/application_tracker.md`
  - Human-readable tracker summary

- `outputs/reports/application_tracker.csv`
  - Machine-readable tracker export

- `application_audit.xlsx`
  - Full Excel audit with ATS and pipeline outputs

## Accounting Persona Split

The accounting layer is intentionally split into two sub-personas.

### `accounting_reporting`

Use for:
- graduate accountant
- junior accountant
- undergraduate accountant
- reporting / compliance / reconciliations / tax roles

Emphasis:
- financial reporting
- reconciliations
- tax compliance
- month-end
- forecast accuracy
- Excel
- accuracy and documentation discipline

### `accounting_systems`

Use for:
- internal accountant
- systems accountant
- finance operations
- AP / AR / invoicing / accounting systems roles

Emphasis:
- AP
- AR
- invoicing
- reconciliations
- month-end
- P/L reporting
- accounting systems
- Xero / MYOB / QuickBooks / ERP / SAP

## Typical CLI Commands

### Install

```powershell
py -m pip install -r requirements.txt
py -m playwright install chromium
```

### Hydrate verified profile facts from Digital Twin

```powershell
py scripts/job_workflow.py hydrate-profile --source-dir "D:\上课\Ai agent\digital twin"
```

### Save browser login session

```powershell
py scripts/job_workflow.py auth --platform seek
py scripts/job_workflow.py auth --platform linkedin
```

### Scrape jobs

```powershell
py scripts/job_workflow.py scrape --platform seek --query "business analyst" --location "Brisbane QLD" --limit 20
py scripts/job_workflow.py scrape --platform linkedin --query "graduate accountant" --location "Melbourne VIC" --limit 20
```

### Run the tailoring pipeline

```powershell
py scripts/job_workflow.py run
py scripts/job_workflow.py run --use-openai
```

### Normalize and export tracker

```powershell
py scripts/job_workflow.py tracker
```

### Assist with applications

```powershell
py scripts/job_workflow.py apply --platform seek --limit 20
py scripts/job_workflow.py apply --platform linkedin --limit 20
```

## How Resume Generation Works

### Input sources

- Base resume text / structure
- `candidate_profile.json`
- `external_profile_facts.json`
- JD content

### Decision flow

1. `jd_parser.py` removes noise
2. `role_classifier.py` classifies the job
3. `resume_analyzer.py` decides what to keep / boost / downplay
4. `profile_filter.py` trims profile facts to a small high-signal set
5. `finance_signal_injector.py` adds evidence-aware finance signals when relevant
6. `resume_generator.py` builds a role-locked resume
7. `ats_scorer.py` scores the result
8. `pipeline_orchestrator.py` regenerates once if below threshold

## Important Safety Rules in Code

- Never delete real job titles, company names, dates, or quantified achievements
- Never merge roles into a generic block
- Never fabricate experience
- Keep resumes ATS-safe and copy-ready
- Keep cover letters separate from ATS analysis
- Use profile evidence levels for finance software and finance operations wording

## Handoff Notes For Copilot

If you continue this project in Copilot, the safest extension points are:

- `scripts/role_classifier.py`
  - add or refine role families

- `scripts/jd_parser.py`
  - improve signal extraction
  - improve noise filtering

- `scripts/resume_generator.py`
  - refine persona-specific title / summary / skills / bullet rewriting

- `scripts/ats_scorer.py`
  - adjust weighting and penalties

- `scripts/finance_signal_injector.py`
  - extend evidence-aware finance signal injection

- `scripts/pipeline_orchestrator.py`
  - adjust refinement loop or add future quality gates

Recommended rule when editing:
- preserve execution automation unless there is a concrete bug
- focus changes on parser, persona, scoring, and orchestration layers

## Current Known Gaps

- Some business analyst roles still underperform when the JD expects stronger stakeholder / workshop / UAT language inside experience bullets
- Some weak accounting jobs still reflect real experience mismatch, not only generation issues
- Some generated summaries can still be made more natural in edge cases
