# Progress / Where We Left Off

Last updated: 2026-07-28

## Live app
https://dash-app-f8ml.onrender.com (Render Web Service, auto-redeploys on push to `main`)

## Status: all planned UI/UX + data-integrity fixes shipped and verified live

The last work block was a full UI/UX audit (user-requested) plus bug fixes, all
committed, pushed, and verified both locally and on Render:

- Fixed a **critical filter bug**: repair ratio was displayed as a formatted
  string ("10.24%"), which made the DataTable's native filter compare
  alphabetically instead of numerically (`>5%` silently excluded `10.24%`).
  Fixed by keeping the underlying value numeric and using
  `dash_table.Format` for display only.
- Fixed **Pareto chart line zigzag** — caused by duplicate `project_no`
  values sharing one x-axis category. Fixed with a unique
  `project_label` (project_no + dimensions).
- Reworked hover tooltips (custom `hovertemplate` + shared `HOVER_STYLE`)
  across all charts — default tooltips were low-contrast and showed raw
  values.
- Simplified Import tab validation UI: success now shows a single green
  line with new/update counts; the 8-item checklist only appears on error
  (previously always shown — user felt it looked amateurish).
- Fixed a contradictory validation message (red "fail" box containing
  green "no errors" text) — "Row count reasonable?" check had no
  corresponding error message.
- Fixed PDF report (A3 landscape) wasting ~40% of page width — now
  computes usable width dynamically, lays out Trend+Amount charts side by
  side, uses PIL to preserve chart aspect ratio. Went from 3 mostly-empty
  pages to 2 well-filled pages. Verified readable, no overlaps.
- Added `dcc.Loading` spinners to the Import tab (Excel + baseline CSV
  upload) — user reported the upload looked "frozen" with no feedback.
- Added retry-with-backoff (`_execute_with_retry`, 2 retries, 0.6s delay)
  around all Supabase calls in `database.py` — a transient
  "JWT issued at future" error was seen twice; in production (gunicorn,
  no debug overlay) this would fail silently, directly explaining the
  "looks frozen" complaint.
- Fixed "Id" column in Pipe Analysis showing "4224.00" instead of "4224"
  (added `id` to `INTEGER_COLUMNS`).

**Verified live on Render (2026-07-28):** uploaded
`Daily Activity Tracking Report 2 - 2026.xlsx` (same-day data, used to
test the update-detection path) — spinner showed during upload, result
correctly reported "0 new, 26 will update existing records for this
date", Id column showed clean integers, no console errors.

## Roadmap (see `/Users/macbookpro/.claude/plans/nested-brewing-spindle.md`)

1. ~~Pipe-Level Analysis~~ — done
2. ~~PDF Report~~ — done
3. **Pipe-level extra charts** — not started. Candidates: pipe-level
   Pareto, repair-category distribution, joint-count distribution,
   outlier scatter.
4. **Group-based charts** — not started. Project Grouping tab already
   lets the user define pipe/machine groups
   (`load_project_group_config`/`upsert_project_group_config` in
   `database.py`), but no chart yet actually uses those saved groups.
5. Module 2 — Data Conversion — placeholder only, scope not yet defined
   with user.
6. Module 3 — Catalog — placeholder only, scope not yet defined with
   user.

## Open question for user
Whether to keep testing with `Daily Activity Tracking Report 2 - 2026.xlsx`
(confirmed content-identical to existing 2026-07-27 data, useful only for
exercising the update-path) or wait for a genuinely different day's data
to test the real "~5-6 pipes change per day" update pattern end to end.

## Conventions / reminders for future work on this repo
- Never copy code from `fuattongethdm/DailyTrackingReport` (the Streamlit
  original, written by someone else) — reference it only for functional
  behavior / data contracts, write fresh implementations here.
- Don't commit real sample `.xlsx` files — already gitignored (`*.xlsx`).
- Ask before pushing/committing when it's not obviously expected.
- After each deploy-affecting change: test locally first, commit with a
  descriptive message, push, then verify on the live Render URL.
