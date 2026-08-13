from __future__ import annotations

import os
import sqlite3
import time
from pathlib import Path
from typing import Any

import pandas as pd

try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass


TABLE_NAME = "repair_rates"
BASELINE_TABLE_NAME = "historical_baselines"
PIPE_TABLE_NAME = "pipe_repair_details"
PROJECT_GROUP_CONFIG_TABLE_NAME = "project_group_configs"
PROJECT_SHEET_LINK_TABLE_NAME = "project_sheet_links"
DB_PATH = Path(os.getenv("REPAIR_DB_PATH", str(Path("data") / "daily_repair_rate.sqlite")))


SCHEMA = """
CREATE TABLE IF NOT EXISTS repair_rates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT NOT NULL,
    production_type TEXT NOT NULL DEFAULT 'Coil',
    project_no TEXT NOT NULL,
    dimensions TEXT NOT NULL,
    qty REAL NOT NULL,
    project_total_pipe_length REAL NOT NULL,
    repaired_pipes_total_length REAL NOT NULL,
    repaired_spiral_length REAL NOT NULL,
    total_repair_amount REAL NOT NULL,
    total_repair_amount_incl_skelp REAL NOT NULL,
    project_status TEXT NOT NULL,
    repair_ratio REAL NOT NULL,
    repair_ratio_incl_skelp REAL NOT NULL,
    UNIQUE(date, project_no, dimensions)
);
"""

BASELINE_SCHEMA = """
CREATE TABLE IF NOT EXISTS historical_baselines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_no TEXT NOT NULL,
    dimensions TEXT NOT NULL,
    repaired_spiral_length REAL NOT NULL,
    total_repair_amount REAL NOT NULL,
    total_repair_amount_incl_skelp REAL NOT NULL,
    project_status TEXT NOT NULL,
    UNIQUE(project_no, dimensions)
);
"""

PIPE_SCHEMA = """
CREATE TABLE IF NOT EXISTS pipe_repair_details (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    first_seen_date TEXT NOT NULL,
    last_updated_date TEXT NOT NULL,
    project_sheet TEXT NOT NULL,
    block_cell TEXT NOT NULL,
    pipe_no INTEGER NOT NULL,
    pipe_length_ft REAL,
    repair_amount REAL NOT NULL,
    repair_ratio REAL NOT NULL,
    repair_count INTEGER,
    repair_category TEXT NOT NULL,
    surface_state TEXT NOT NULL,
    UNIQUE(project_sheet, block_cell)
);
"""

PROJECT_GROUP_CONFIG_SCHEMA = """
CREATE TABLE IF NOT EXISTS project_group_configs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_sheet TEXT NOT NULL,
    project_no TEXT NOT NULL,
    dimensions TEXT NOT NULL,
    pipe_groups TEXT NOT NULL DEFAULT '',
    machine_groups TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(project_sheet, project_no, dimensions)
);
"""

PROJECT_SHEET_LINK_SCHEMA = """
CREATE TABLE IF NOT EXISTS project_sheet_links (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_sheet TEXT NOT NULL,
    project_no TEXT,
    dimensions TEXT,
    status TEXT NOT NULL DEFAULT 'confirmed',
    match_score REAL,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(project_sheet)
);
"""


UPSERT_SQL = """
INSERT INTO repair_rates (
    date, production_type, project_no, dimensions, qty, project_total_pipe_length,
    repaired_pipes_total_length, repaired_spiral_length, total_repair_amount,
    total_repair_amount_incl_skelp, project_status, repair_ratio,
    repair_ratio_incl_skelp
) VALUES (
    :date, :production_type, :project_no, :dimensions, :qty, :project_total_pipe_length,
    :repaired_pipes_total_length, :repaired_spiral_length, :total_repair_amount,
    :total_repair_amount_incl_skelp, :project_status, :repair_ratio,
    :repair_ratio_incl_skelp
)
ON CONFLICT(date, project_no, dimensions) DO UPDATE SET
    production_type = excluded.production_type,
    qty = excluded.qty,
    project_total_pipe_length = excluded.project_total_pipe_length,
    repaired_pipes_total_length = excluded.repaired_pipes_total_length,
    repaired_spiral_length = excluded.repaired_spiral_length,
    total_repair_amount = excluded.total_repair_amount,
    total_repair_amount_incl_skelp = excluded.total_repair_amount_incl_skelp,
    project_status = excluded.project_status,
    repair_ratio = excluded.repair_ratio,
    repair_ratio_incl_skelp = excluded.repair_ratio_incl_skelp;
"""

BASELINE_UPSERT_SQL = """
INSERT INTO historical_baselines (
    project_no, dimensions, repaired_spiral_length, total_repair_amount,
    total_repair_amount_incl_skelp, project_status
) VALUES (
    :project_no, :dimensions, :repaired_spiral_length, :total_repair_amount,
    :total_repair_amount_incl_skelp, :project_status
)
ON CONFLICT(project_no, dimensions) DO UPDATE SET
    repaired_spiral_length = excluded.repaired_spiral_length,
    total_repair_amount = excluded.total_repair_amount,
    total_repair_amount_incl_skelp = excluded.total_repair_amount_incl_skelp,
    project_status = excluded.project_status;
"""

PIPE_UPSERT_SQL = """
INSERT INTO pipe_repair_details (
    first_seen_date, last_updated_date, project_sheet, block_cell, pipe_no,
    pipe_length_ft, repair_amount, repair_ratio, repair_count, repair_category,
    surface_state
) VALUES (
    :first_seen_date, :last_updated_date, :project_sheet, :block_cell, :pipe_no,
    :pipe_length_ft, :repair_amount, :repair_ratio, :repair_count, :repair_category,
    :surface_state
)
ON CONFLICT(project_sheet, block_cell) DO UPDATE SET
    last_updated_date = excluded.last_updated_date,
    pipe_no = excluded.pipe_no,
    pipe_length_ft = excluded.pipe_length_ft,
    repair_amount = excluded.repair_amount,
    repair_ratio = excluded.repair_ratio,
    repair_count = excluded.repair_count,
    repair_category = excluded.repair_category,
    surface_state = excluded.surface_state;
"""

PROJECT_GROUP_CONFIG_UPSERT_SQL = """
INSERT INTO project_group_configs (
    project_sheet, project_no, dimensions, pipe_groups, machine_groups, updated_at
) VALUES (
    :project_sheet, :project_no, :dimensions, :pipe_groups, :machine_groups, CURRENT_TIMESTAMP
)
ON CONFLICT(project_sheet, project_no, dimensions) DO UPDATE SET
    pipe_groups = excluded.pipe_groups,
    machine_groups = excluded.machine_groups,
    updated_at = CURRENT_TIMESTAMP;
"""


def _get_config_value(name: str) -> str | None:
    # Read from environment variables only (.env or a real env var).
    return os.getenv(name)


def get_backend_name() -> str:
    if _get_config_value("SUPABASE_URL") and _get_config_value("SUPABASE_KEY"):
        return "supabase"
    return "sqlite"


def get_connection(db_path: Path = DB_PATH) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


_supabase_client = None


def _get_supabase_client():
    global _supabase_client
    if _supabase_client is not None:
        return _supabase_client

    try:
        from supabase import create_client
    except ImportError as exc:
        raise RuntimeError("Supabase backend selected but 'supabase' package is not installed.") from exc

    url = _get_config_value("SUPABASE_URL")
    key = _get_config_value("SUPABASE_KEY")
    if not url or not key:
        raise RuntimeError("SUPABASE_URL and SUPABASE_KEY must be set for Supabase backend.")

    for proxy_var in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
        if os.getenv(proxy_var, "").strip().lower() == "http://127.0.0.1:9":
            os.environ.pop(proxy_var, None)

    _supabase_client = create_client(url, key)
    return _supabase_client


def _execute_with_retry(query, retries: int = 2, delay: float = 0.6):
    """Run a Supabase query, retrying transient failures a couple of times.

    Supabase occasionally rejects a request with a "JWT issued at future"
    auth error with no persistent cause (the local clock is correct) —
    without a retry, that single blip fails the whole callback, and in
    production (no dev error overlay) it just looks like the app froze.
    """
    last_exc: Exception | None = None
    for attempt in range(retries + 1):
        try:
            return query.execute()
        except Exception as exc:  # noqa: BLE001 - deliberately broad, retried below
            last_exc = exc
            if attempt < retries:
                time.sleep(delay)
    raise last_exc


def _fetch_supabase_table(table_name: str, order_columns: tuple[str, ...], page_size: int = 1000) -> list[dict[str, Any]]:
    client = _get_supabase_client()
    rows: list[dict[str, Any]] = []
    for start in range(0, 100_000, page_size):
        query = client.table(table_name).select("*")
        for column in order_columns:
            query = query.order(column)
        response = _execute_with_retry(query.range(start, start + page_size - 1))
        rows.extend(response.data)
        if len(response.data) < page_size:
            break
    return rows


# ---------------------------------------------------------------------------
# In-process cache for the full-table loaders below. The Dashboard tab fires
# many independent callbacks on every page load, and several of them load
# the exact same table (e.g. load_master_data() is called from 8 different
# callbacks) — without this, each one pays its own full Supabase round trip.
# Keyed by table name; a short TTL is a safety net for data changed outside
# this app's own write path, on top of explicit invalidation on every write.
# ---------------------------------------------------------------------------

_CACHE: dict[str, tuple[float, pd.DataFrame]] = {}
_CACHE_TTL_SECONDS = 120


def _cache_get(key: str) -> pd.DataFrame | None:
    entry = _CACHE.get(key)
    if entry is None:
        return None
    cached_at, df = entry
    if time.time() - cached_at > _CACHE_TTL_SECONDS:
        return None
    return df.copy()


def _cache_set(key: str, df: pd.DataFrame) -> None:
    _CACHE[key] = (time.time(), df.copy())


def _cache_invalidate(*keys: str) -> None:
    for key in keys:
        _CACHE.pop(key, None)


def init_db(conn: sqlite3.Connection | None = None) -> None:
    if get_backend_name() == "supabase":
        return

    should_close = conn is None
    conn = conn or get_connection()
    conn.execute(SCHEMA)
    conn.execute(BASELINE_SCHEMA)
    conn.execute(PIPE_SCHEMA)
    conn.execute(PROJECT_GROUP_CONFIG_SCHEMA)
    conn.execute(PROJECT_SHEET_LINK_SCHEMA)
    columns = {row["name"] for row in conn.execute("PRAGMA table_info(repair_rates)").fetchall()}
    if "production_type" not in columns:
        conn.execute("ALTER TABLE repair_rates ADD COLUMN production_type TEXT NOT NULL DEFAULT 'Coil'")
    conn.commit()
    if should_close:
        conn.close()


def load_project_group_config(
    project_sheet: str,
    project_no: str,
    dimensions: str,
    conn: sqlite3.Connection | None = None,
) -> dict[str, str] | None:
    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        try:
            response = _execute_with_retry(
                client.table(PROJECT_GROUP_CONFIG_TABLE_NAME)
                .select("pipe_groups,machine_groups")
                .eq("project_sheet", project_sheet)
                .eq("project_no", project_no)
                .eq("dimensions", dimensions)
                .limit(1)
            )
        except Exception as exc:
            if PROJECT_GROUP_CONFIG_TABLE_NAME in str(exc) and ("PGRST205" in str(exc) or "schema cache" in str(exc)):
                return None
            raise
        if not response.data:
            return None
        row = response.data[0]
        return {
            "pipe_groups": row.get("pipe_groups") or "",
            "machine_groups": row.get("machine_groups") or "",
        }

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    row = conn.execute(
        """
        SELECT pipe_groups, machine_groups
        FROM project_group_configs
        WHERE project_sheet = ? AND project_no = ? AND dimensions = ?
        LIMIT 1
        """,
        (project_sheet, project_no, dimensions),
    ).fetchone()
    if should_close:
        conn.close()
    if row is None:
        return None
    return {"pipe_groups": row["pipe_groups"] or "", "machine_groups": row["machine_groups"] or ""}


def upsert_project_group_config(
    project_sheet: str,
    project_no: str,
    dimensions: str,
    pipe_groups: str,
    machine_groups: str,
    conn: sqlite3.Connection | None = None,
) -> int:
    record = {
        "project_sheet": project_sheet,
        "project_no": project_no,
        "dimensions": dimensions,
        "pipe_groups": pipe_groups.strip(),
        "machine_groups": machine_groups.strip(),
    }

    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        _execute_with_retry(
            client.table(PROJECT_GROUP_CONFIG_TABLE_NAME).upsert(
                record,
                on_conflict="project_sheet,project_no,dimensions",
            )
        )
        return 1

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    conn.execute(PROJECT_GROUP_CONFIG_UPSERT_SQL, record)
    conn.commit()
    if should_close:
        conn.close()
    return 1


def load_project_sheet_links(conn: sqlite3.Connection | None = None) -> pd.DataFrame:
    """Every project_sheet -> project_no/dimensions mapping confirmed via the
    "Project Mapping" review tool (feature/project-sheet-mapping-tool branch).
    Read-only here — mappings are written from that branch, not this one."""
    cached = _cache_get(PROJECT_SHEET_LINK_TABLE_NAME)
    if cached is not None:
        return cached

    empty = pd.DataFrame(columns=["project_sheet", "project_no", "dimensions", "status", "match_score"])
    if get_backend_name() == "supabase":
        try:
            rows = _fetch_supabase_table(PROJECT_SHEET_LINK_TABLE_NAME, ("project_sheet",))
        except Exception as exc:
            if PROJECT_SHEET_LINK_TABLE_NAME in str(exc) and ("PGRST205" in str(exc) or "schema cache" in str(exc)):
                _cache_set(PROJECT_SHEET_LINK_TABLE_NAME, empty)
                return empty
            raise
        df = pd.DataFrame(rows) if rows else empty
        _cache_set(PROJECT_SHEET_LINK_TABLE_NAME, df)
        return df

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    df = pd.read_sql_query("SELECT * FROM project_sheet_links ORDER BY project_sheet", conn)
    if should_close:
        conn.close()
    _cache_set(PROJECT_SHEET_LINK_TABLE_NAME, df)
    return df


def _records_from_df(df: pd.DataFrame) -> list[dict[str, Any]]:
    write_df = df.drop(columns=["excel_row", "id"], errors="ignore").copy()
    if "production_type" not in write_df.columns:
        write_df["production_type"] = "Coil"
    write_df["production_type"] = write_df["production_type"].fillna("Coil")
    write_df["date"] = pd.to_datetime(write_df["date"]).dt.strftime("%Y-%m-%d")
    records = write_df.where(pd.notna(write_df), None).to_dict(orient="records")
    return records


def _baseline_records_from_df(df: pd.DataFrame) -> list[dict[str, Any]]:
    write_df = df[
        [
            "project_no",
            "dimensions",
            "repaired_spiral_length",
            "total_repair_amount",
            "total_repair_amount_incl_skelp",
            "project_status",
        ]
    ].copy()
    return write_df.where(pd.notna(write_df), None).to_dict(orient="records")


def _infer_baseline_production_type(project_no) -> str:
    """Same naming convention the daily-table parser's Plate section uses:
    plate jobs carry "(PT)" or "Plate" in the project number. Archive rows
    don't carry an explicit type column, so this is inferred from the name."""
    lowered = str(project_no).lower()
    return "Plate" if ("(pt)" in lowered or "plate" in lowered) else "Coil"


def upsert_historical_baselines(df: pd.DataFrame, conn: sqlite3.Connection | None = None) -> int:
    if df.empty:
        return 0
    records = _baseline_records_from_df(df)

    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        _execute_with_retry(client.table(BASELINE_TABLE_NAME).upsert(records, on_conflict="project_no,dimensions"))
        _cache_invalidate(BASELINE_TABLE_NAME)
        return len(records)

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    conn.executemany(BASELINE_UPSERT_SQL, records)
    conn.commit()
    if should_close:
        conn.close()
    _cache_invalidate(BASELINE_TABLE_NAME)
    return len(records)


def load_historical_baselines(conn: sqlite3.Connection | None = None) -> pd.DataFrame:
    cached = _cache_get(BASELINE_TABLE_NAME)
    if cached is not None:
        return cached

    if get_backend_name() == "supabase":
        try:
            rows = _fetch_supabase_table(BASELINE_TABLE_NAME, ("project_no", "dimensions"))
        except Exception as exc:
            if "historical_baselines" in str(exc) and ("PGRST205" in str(exc) or "schema cache" in str(exc)):
                _cache_set(BASELINE_TABLE_NAME, pd.DataFrame())
                return pd.DataFrame()
            raise
        df = pd.DataFrame(rows)
    else:
        should_close = conn is None
        conn = conn or get_connection()
        init_db(conn)
        df = pd.read_sql_query("SELECT * FROM historical_baselines ORDER BY project_no, dimensions", conn)
        if should_close:
            conn.close()

    if not df.empty:
        df["production_type"] = df["project_no"].map(_infer_baseline_production_type)
    _cache_set(BASELINE_TABLE_NAME, df)
    return df


def _pipe_records_from_df(df: pd.DataFrame) -> list[dict[str, Any]]:
    """Build upsert-ready records from a freshly-parsed pipe_df (which only
    carries a single ``date`` — the day of this upload). Both
    ``first_seen_date`` and ``last_updated_date`` default to that date here;
    for a pipe that already exists, the caller (upsert_pipe_repair_details)
    overrides ``first_seen_date`` with its real, original value before
    writing, so a pipe's repair date never changes on re-upload."""
    write_df = df.drop(columns=["id"], errors="ignore").copy()
    upload_date = pd.to_datetime(write_df["date"]).dt.strftime("%Y-%m-%d")
    write_df["first_seen_date"] = upload_date
    write_df["last_updated_date"] = upload_date
    write_df = write_df.drop(columns=["date"])
    records = write_df.astype(object).where(pd.notna(write_df), None).to_dict(orient="records")
    for record in records:
        record["pipe_no"] = int(record["pipe_no"])
        if record.get("repair_count") is not None:
            record["repair_count"] = int(record["repair_count"])
    return records


def get_existing_keys(df: pd.DataFrame, conn: sqlite3.Connection | None = None) -> set[tuple[str, str, str]]:
    if df.empty:
        return set()

    dates = sorted(set(pd.to_datetime(df["date"]).dt.strftime("%Y-%m-%d")))

    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        response = _execute_with_retry(
            client.table(TABLE_NAME).select("date,project_no,dimensions").in_("date", dates)
        )
        return {(row["date"], row["project_no"], row["dimensions"]) for row in response.data}

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    placeholders = ",".join("?" for _ in dates)
    rows = conn.execute(
        f"SELECT date, project_no, dimensions FROM repair_rates WHERE date IN ({placeholders})",
        dates,
    ).fetchall()
    if should_close:
        conn.close()
    return {(row["date"], row["project_no"], row["dimensions"]) for row in rows}


def upsert_repair_rates(df: pd.DataFrame, conn: sqlite3.Connection | None = None) -> int:
    records = _records_from_df(df)

    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        _execute_with_retry(client.table(TABLE_NAME).upsert(records, on_conflict="date,project_no,dimensions"))
        _cache_invalidate(TABLE_NAME)
        return len(records)

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    conn.executemany(UPSERT_SQL, records)
    conn.commit()
    if should_close:
        conn.close()
    _cache_invalidate(TABLE_NAME)
    return len(records)


def load_master_data(conn: sqlite3.Connection | None = None) -> pd.DataFrame:
    cached = _cache_get(TABLE_NAME)
    if cached is not None:
        return cached

    if get_backend_name() == "supabase":
        df = pd.DataFrame(_fetch_supabase_table(TABLE_NAME, ("date", "project_no", "dimensions")))
    else:
        should_close = conn is None
        conn = conn or get_connection()
        init_db(conn)
        df = pd.read_sql_query("SELECT * FROM repair_rates ORDER BY date, project_no, dimensions", conn)
        if should_close:
            conn.close()

    if not df.empty:
        df["date"] = pd.to_datetime(df["date"])
        if "production_type" not in df.columns:
            df["production_type"] = "Coil"
        df["production_type"] = df["production_type"].fillna("Coil")
    _cache_set(TABLE_NAME, df)
    return df


def _existing_pipe_first_seen() -> dict[tuple[str, str], str]:
    """(project_sheet, block_cell) -> first_seen_date for every pipe already
    on file — the whole table, no date filter, since the table is now
    bounded to one row per physical pipe rather than one per pipe per day.
    Supabase-only: PostgREST's upsert can't "update every column except
    one," so the caller uses this to preserve a pipe's original
    first_seen_date in Python before re-upserting it. SQLite's
    ON CONFLICT ... DO UPDATE already leaves first_seen_date untouched
    natively, so it doesn't need this lookup."""
    client = _get_supabase_client()
    rows: list[dict[str, Any]] = []
    page_size = 1000
    try:
        for start in range(0, 100_000, page_size):
            response = _execute_with_retry(
                client.table(PIPE_TABLE_NAME)
                .select("project_sheet,block_cell,first_seen_date")
                .order("project_sheet")
                .order("block_cell")
                .range(start, start + page_size - 1)
            )
            rows.extend(response.data)
            if len(response.data) < page_size:
                break
    except Exception as exc:
        if PIPE_TABLE_NAME in str(exc) and ("PGRST205" in str(exc) or "schema cache" in str(exc)):
            return {}
        raise
    return {(row["project_sheet"], row["block_cell"]): row["first_seen_date"] for row in rows}


def upsert_pipe_repair_details(df: pd.DataFrame, conn: sqlite3.Connection | None = None) -> int:
    if df.empty:
        return 0
    records = _pipe_records_from_df(df)

    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        existing_first_seen = _existing_pipe_first_seen()
        for record in records:
            key = (record["project_sheet"], record["block_cell"])
            if key in existing_first_seen:
                record["first_seen_date"] = existing_first_seen[key]
        for start in range(0, len(records), 500):
            batch = records[start : start + 500]
            _execute_with_retry(
                client.table(PIPE_TABLE_NAME).upsert(
                    batch,
                    on_conflict="project_sheet,block_cell",
                )
            )
        _cache_invalidate(PIPE_TABLE_NAME)
        return len(records)

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    conn.executemany(PIPE_UPSERT_SQL, records)
    conn.commit()
    if should_close:
        conn.close()
    _cache_invalidate(PIPE_TABLE_NAME)
    return len(records)


def load_pipe_repair_details(conn: sqlite3.Connection | None = None) -> pd.DataFrame:
    cached = _cache_get(PIPE_TABLE_NAME)
    if cached is not None:
        return cached

    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        try:
            rows: list[dict[str, Any]] = []
            page_size = 1000
            for start in range(0, 100_000, page_size):
                response = _execute_with_retry(
                    client.table(PIPE_TABLE_NAME)
                    .select("*")
                    # order() must be a fully-unique key for .range() paging to be
                    # stable — first_seen_date alone has heavy ties (many pipes
                    # share a date), which made range() return some rows twice
                    # and skip others. (project_sheet, block_cell) matches the
                    # table's own UNIQUE constraint.
                    .order("project_sheet")
                    .order("block_cell")
                    .range(start, start + page_size - 1)
                )
                rows.extend(response.data)
                if len(response.data) < page_size:
                    break
        except Exception as exc:
            if PIPE_TABLE_NAME in str(exc) and ("PGRST205" in str(exc) or "schema cache" in str(exc)):
                empty = pd.DataFrame()
                _cache_set(PIPE_TABLE_NAME, empty)
                return empty
            raise
        df = pd.DataFrame(rows)
    else:
        should_close = conn is None
        conn = conn or get_connection()
        init_db(conn)
        df = pd.read_sql_query(
            "SELECT * FROM pipe_repair_details ORDER BY first_seen_date, project_sheet, pipe_no", conn
        )
        if should_close:
            conn.close()

    df = _normalize_pipe_repair_details(df)
    _cache_set(PIPE_TABLE_NAME, df)
    return df


def load_pipe_repair_details_for_date(selected_date, conn: sqlite3.Connection | None = None) -> pd.DataFrame:
    """Pipes whose first_seen_date (i.e. the day they were repaired) is
    ``selected_date``. Not called anywhere yet, kept for future callers."""
    date_text = pd.to_datetime(selected_date).date().isoformat()
    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        try:
            rows: list[dict[str, Any]] = []
            page_size = 1000
            for start in range(0, 100_000, page_size):
                response = _execute_with_retry(
                    client.table(PIPE_TABLE_NAME)
                    .select("*")
                    .eq("first_seen_date", date_text)
                    .order("project_sheet")
                    .order("block_cell")
                    .range(start, start + page_size - 1)
                )
                rows.extend(response.data)
                if len(response.data) < page_size:
                    break
        except Exception as exc:
            if PIPE_TABLE_NAME in str(exc) and ("PGRST205" in str(exc) or "schema cache" in str(exc)):
                return pd.DataFrame()
            raise
        return _normalize_pipe_repair_details(pd.DataFrame(rows))

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    df = pd.read_sql_query(
        "SELECT * FROM pipe_repair_details WHERE first_seen_date = ? ORDER BY project_sheet, pipe_no",
        conn,
        params=(date_text,),
    )
    if should_close:
        conn.close()
    return _normalize_pipe_repair_details(df)


def _normalize_pipe_repair_details(df: pd.DataFrame) -> pd.DataFrame:
    if not df.empty:
        df["first_seen_date"] = pd.to_datetime(df["first_seen_date"])
        df["last_updated_date"] = pd.to_datetime(df["last_updated_date"])
        placeholder_mask = pd.to_numeric(df["repair_amount"], errors="coerce").le(0.0001)
        df.loc[placeholder_mask, ["repair_amount", "repair_ratio"]] = 0.0
        sort_columns = [
            column for column in ["first_seen_date", "project_sheet", "pipe_no"] if column in df.columns
        ]
        if sort_columns:
            df = df.sort_values(sort_columns).reset_index(drop=True)
    return df


