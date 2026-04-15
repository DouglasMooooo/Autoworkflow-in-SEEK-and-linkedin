# Career-Ops Integration Plan

This repo should stay centered on your current strengths:

- customized resume generation from your own source templates
- browser-driven Seek and LinkedIn flows using your existing Edge login state
- OpenAI / Codex-compatible execution

Instead of replacing that foundation, the best parts of `career-ops` should be adopted as a support layer.

## What To Keep From Your Current Project

- Python-first workflow in `scripts/job_workflow.py`
- `jobs/jobs.csv` as a simple inbox
- `browser_profiles/` for persistent authenticated sessions
- `candidate_profile.json` and your existing resume templates
- direct `run`, `scrape`, `apply` commands

## What To Borrow From Career-Ops

- strict separation of user data vs system logic
- tracker-first thinking, not just one-off output generation
- deduplication and status normalization
- merge-safe updates so reruns do not wipe application history
- markdown tracker views for quick review
- stronger pipeline operations before full automation

## What Was Integrated In This Pass

- Audit rows are now merged with existing history instead of being blindly overwritten on every `run`
- Statuses are normalized into a canonical set
- Duplicate roles are collapsed by URL or company/title identity
- Better scoring rows can refresh resume and cover-letter artifacts without losing historical status
- Tracker exports are now generated automatically:
  - `outputs/reports/application_tracker.csv`
  - `outputs/reports/application_tracker.md`
- Added `py scripts/job_workflow.py tracker` to rebuild and summarize the tracker

## Recommended Next Upgrades

1. Add a company portal scanner layer for Greenhouse, Lever, and Ashby jobs.
2. Add a richer pipeline inbox separate from `jobs/jobs.csv`, so scraped roles and manually added roles can coexist cleanly.
3. Split the current monolithic Python file into modules:
   - tracker
   - resume generation
   - scraping
   - browser apply assistant
4. Add a profile-driven fit scoring model that includes:
   - location fit
   - visa fit
   - seniority fit
   - salary fit
   - role archetype fit
5. Add batch evaluation support using Codex/OpenAI models rather than Claude-specific tooling.

## Codex-Friendly Mapping

The original project assumes Claude-centric agent workflows in several places. In this repo, that should map as:

- Claude mode prompts -> OpenAI Responses API prompts
- Claude slash-command mindset -> Python CLI subcommands
- Claude tracker hygiene scripts -> Python tracker merge/dedup helpers
- Claude-first customization loop -> Codex-assisted repo evolution

## Practical Rule

Use `career-ops` as an idea library and operating model, not as a drop-in replacement.

That lets this project stay optimized for your real-world application flow while still getting much stronger operational discipline.
