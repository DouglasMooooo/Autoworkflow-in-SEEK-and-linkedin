# Data Contract

This project now separates personal assets from reusable system logic so the workflow can absorb ideas from `career-ops` without overwriting your own materials.

## User Layer

These files are your source of truth and should not be replaced by upstream integrations:

| Path | Purpose |
|---|---|
| `config/candidate_profile.json` | Your personal identity, work rights, availability, achievements |
| `config/external_profile_facts.json` | Supplemental verified facts imported from your Digital Twin profile project |
| `config/portals.yaml` | Your target companies and title filters for portal scanning |
| `config/settings.yaml` | Your preferred models, templates, and known skills |
| `resumes/*` | Your base resume templates and formatted resume assets |
| `jobs/jobs.csv` | Your current job inbox and manually curated leads |
| `browser_profiles/*` | Your Seek and LinkedIn login state |
| `.env` | Your API keys and model overrides |

## System Layer

These files contain reusable automation logic and can evolve as we merge in stronger ideas:

| Path | Purpose |
|---|---|
| `scripts/job_workflow.py` | Main pipeline, scraping, resume generation, and apply assistant |
| `outputs/reports/application_tracker.csv` | Normalized tracker export |
| `outputs/reports/application_tracker.md` | Human-readable tracker dashboard |
| `outputs/reports/pipeline.md` | Scan pipeline inbox (pending/processed roles) |
| `outputs/reports/scan_history.tsv` | Historical portal scan dedupe ledger |
| `application_audit.xlsx` | Working audit sheet, now merge-safe and deduplicated |
| `README.md`, `README_zh.md` | Project instructions |

## Merge Direction

The intended integration pattern is:

1. Keep your execution workflow and browser-authenticated applying flow.
2. Borrow operational ideas from `career-ops`, especially tracker hygiene, deduplication, status normalization, and stronger pipeline reporting.
3. Swap provider-specific assumptions freely.
   `career-ops` may assume Claude-first flows, while this project should stay OpenAI/Codex-friendly.

## Rule

If a file contains your profile, resume source, job targets, or login state, treat it as user-owned.

If a file contains orchestration logic, reporting, or pipeline helpers, treat it as system-owned and safe to improve.
