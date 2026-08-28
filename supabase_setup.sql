-- Factory Tracking System (Dash) — Supabase schema setup
-- Run this entire file once in Supabase Dashboard > SQL Editor.
-- The schema uses the exact same fields as the SQLite schema in dash_app/database.py.

create table if not exists repair_rates (
    id bigint generated always as identity primary key,
    date date not null,
    production_type text not null default 'Coil',
    project_no text not null,
    dimensions text not null,
    qty numeric not null,
    project_total_pipe_length numeric not null,
    repaired_pipes_total_length numeric not null,
    repaired_spiral_length numeric not null,
    total_repair_amount numeric not null,
    total_repair_amount_incl_skelp numeric not null,
    project_status text not null,
    repair_ratio numeric not null,
    repair_ratio_incl_skelp numeric not null,
    unique (date, project_no, dimensions)
);

-- One row per physical pipe (project_sheet + block_cell identifies the
-- pipe's stable Excel repair-record slot), not one row per pipe per day —
-- first_seen_date is set once and never overwritten on re-upload;
-- last_updated_date refreshes on every re-upload. A pipe now has two
-- lifecycle stages, tracked via status:
--   'Produced' — cut, but no repair data yet (repair_amount/repair_ratio/
--                surface_state are null; first_seen_date is the day the
--                pipe was first recorded, i.e. its production/cut date).
--   'Repaired' — repair_amount/repair_ratio present. repaired_date is set
--                the first time this happens and never overwritten after
--                (mirrors first_seen_date's own "set once" behavior) — see
--                database.py:upsert_pipe_repair_details for exactly how a
--                pipe transitions from Produced to Repaired, possibly on
--                the very same day it's first seen (an urgent/priority
--                pipe can be produced and repaired before the same day's
--                upload, in which case it skips the Produced stage
--                entirely and is recorded as Repaired from day one).
create table if not exists pipe_repair_details (
    id bigint generated always as identity primary key,
    first_seen_date date not null,
    last_updated_date date not null,
    project_sheet text not null,
    block_cell text not null,
    pipe_no integer not null,
    pipe_length_ft numeric,
    repair_amount numeric,
    repair_ratio numeric,
    repair_count integer,
    repair_category text not null,
    surface_state text,
    status text not null default 'Repaired',
    repaired_date date,
    unique (project_sheet, block_cell)
);

create table if not exists project_group_configs (
    id bigint generated always as identity primary key,
    project_sheet text not null,
    project_no text not null,
    dimensions text not null,
    pipe_groups text not null default '',
    machine_groups text not null default '',
    updated_at timestamptz not null default now(),
    unique (project_sheet, project_no, dimensions)
);

-- RLS stays enabled by default. The app backend (Dash/Python) must connect
-- with the "service_role" key — service_role bypasses RLS automatically,
-- so no extra policies are needed. Set SUPABASE_KEY in .env to the
-- "service_role" secret key (not the anon/public key).
alter table repair_rates enable row level security;
alter table pipe_repair_details enable row level security;
alter table project_group_configs enable row level security;

-- ---------------------------------------------------------------------------
-- MIGRATION for an existing pipe_repair_details table (the `create table if
-- not exists` above only applies to a brand-new project). Run this block
-- once, by hand, when deploying the "track produced-but-unrepaired pipes"
-- feature to a database that already has pipe_repair_details with data in
-- it. Safe to run more than once (every statement is idempotent).
-- ---------------------------------------------------------------------------
-- alter table pipe_repair_details alter column repair_amount drop not null;
-- alter table pipe_repair_details alter column repair_ratio drop not null;
-- alter table pipe_repair_details alter column surface_state drop not null;
-- alter table pipe_repair_details add column if not exists status text not null default 'Repaired';
-- alter table pipe_repair_details add column if not exists repaired_date date;
-- -- Every row that already exists today is, by definition, a repaired pipe
-- -- (this feature is what introduces the "Produced" stage) — the column
-- -- default above already covers this, this update is only needed for
-- -- repaired_date, backfilled from first_seen_date (the closest available
-- -- approximation for pipes repaired before this feature existed).
-- update pipe_repair_details set repaired_date = first_seen_date where repaired_date is null;

-- ---------------------------------------------------------------------------
-- MIGRATION for the existing project_sheet_links table (not itself defined
-- in this file yet — it predates this script; this only adds one column to
-- whatever's already live). Run once, by hand, in the SQL Editor. Adds a
-- per-project "are first_seen_date/repaired_date trustworthy" flag, so a
-- project with known-bad dates (e.g. a gap in daily uploads that made the
-- app misattribute when a pipe was really produced/repaired) can be marked
-- unreliable and excluded from cross-project date-based aggregates
-- (Newest Produced/Repaired Pipes, the Backlog Trend / Daily Repair Amount
-- charts) without touching any single-project view — see
-- pages/home.py:_pipe_sheet_label_map's trusted_only parameter.
-- ---------------------------------------------------------------------------
-- alter table project_sheet_links add column if not exists dates_reliable boolean not null default true;
-- update project_sheet_links set dates_reliable = false where project_sheet = '06-131';
