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
-- first_seen_date is set once (the day the pipe was repaired) and never
-- overwritten on re-upload; last_updated_date refreshes on every re-upload.
create table if not exists pipe_repair_details (
    id bigint generated always as identity primary key,
    first_seen_date date not null,
    last_updated_date date not null,
    project_sheet text not null,
    block_cell text not null,
    pipe_no integer not null,
    pipe_length_ft numeric,
    repair_amount numeric not null,
    repair_ratio numeric not null,
    repair_count integer,
    repair_category text not null,
    surface_state text not null,
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
