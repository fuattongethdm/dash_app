from __future__ import annotations

import os
import sqlite3
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
    date TEXT NOT NULL,
    project_sheet TEXT NOT NULL,
    block_cell TEXT NOT NULL,
    pipe_no INTEGER NOT NULL,
    pipe_length_ft REAL,
    repair_amount REAL NOT NULL,
    repair_ratio REAL NOT NULL,
    repair_count INTEGER,
    repair_category TEXT NOT NULL,
    surface_state TEXT NOT NULL,
    UNIQUE(date, project_sheet, block_cell)
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
    date, project_sheet, block_cell, pipe_no, pipe_length_ft, repair_amount,
    repair_ratio, repair_count, repair_category, surface_state
) VALUES (
    :date, :project_sheet, :block_cell, :pipe_no, :pipe_length_ft, :repair_amount,
    :repair_ratio, :repair_count, :repair_category, :surface_state
)
ON CONFLICT(date, project_sheet, block_cell) DO UPDATE SET
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
    # NOT: Streamlit secrets kaldırıldı (Dash'e geçiş). Artık sadece ortam
    # değişkenlerinden (.env veya gerçek env var) okunuyor.
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


def _get_supabase_client():
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

    return create_client(url, key)


def _fetch_supabase_table(table_name: str, order_columns: tuple[str, ...], page_size: int = 1000) -> list[dict[str, Any]]:
    client = _get_supabase_client()
    rows: list[dict[str, Any]] = []
    for start in range(0, 100_000, page_size):
        query = client.table(table_name).select("*")
        for column in order_columns:
            query = query.order(column)
        response = query.range(start, start + page_size - 1).execute()
        rows.extend(response.data)
        if len(response.data) < page_size:
            break
    return rows


def init_db(conn: sqlite3.Connection | None = None) -> None:
    if get_backend_name() == "supabase":
        return

    should_close = conn is None
    conn = conn or get_connection()
    conn.execute(SCHEMA)
    conn.execute(BASELINE_SCHEMA)
    conn.execute(PIPE_SCHEMA)
    conn.execute(PROJECT_GROUP_CONFIG_SCHEMA)
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
            response = (
                client.table(PROJECT_GROUP_CONFIG_TABLE_NAME)
                .select("pipe_groups,machine_groups")
                .eq("project_sheet", project_sheet)
                .eq("project_no", project_no)
                .eq("dimensions", dimensions)
                .limit(1)
                .execute()
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
        client.table(PROJECT_GROUP_CONFIG_TABLE_NAME).upsert(
            record,
            on_conflict="project_sheet,project_no,dimensions",
        ).execute()
        return 1

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    conn.execute(PROJECT_GROUP_CONFIG_UPSERT_SQL, record)
    conn.commit()
    if should_close:
        conn.close()
    return 1


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


def _pipe_records_from_df(df: pd.DataFrame) -> list[dict[str, Any]]:
    write_df = df.drop(columns=["id"], errors="ignore").copy()
    write_df["date"] = pd.to_datetime(write_df["date"]).dt.strftime("%Y-%m-%d")
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
        response = (
            client.table(TABLE_NAME)
            .select("date,project_no,dimensions")
            .in_("date", dates)
            .execute()
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
        client.table(TABLE_NAME).upsert(records, on_conflict="date,project_no,dimensions").execute()
        return len(records)

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    conn.executemany(UPSERT_SQL, records)
    conn.commit()
    if should_close:
        conn.close()
    return len(records)


def load_master_data(conn: sqlite3.Connection | None = None) -> pd.DataFrame:
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
    return df


def upsert_pipe_repair_details(df: pd.DataFrame, conn: sqlite3.Connection | None = None) -> int:
    if df.empty:
        return 0
    records = _pipe_records_from_df(df)

    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        for start in range(0, len(records), 500):
            batch = records[start : start + 500]
            client.table(PIPE_TABLE_NAME).upsert(
                batch,
                on_conflict="date,project_sheet,block_cell",
            ).execute()
        return len(records)

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    conn.executemany(PIPE_UPSERT_SQL, records)
    conn.commit()
    if should_close:
        conn.close()
    return len(records)


def load_pipe_repair_details(conn: sqlite3.Connection | None = None) -> pd.DataFrame:
    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        try:
            rows: list[dict[str, Any]] = []
            page_size = 1000
            for start in range(0, 100_000, page_size):
                response = (
                    client.table(PIPE_TABLE_NAME)
                    .select("*")
                    .order("date")
                    .range(start, start + page_size - 1)
                    .execute()
                )
                rows.extend(response.data)
                if len(response.data) < page_size:
                    break
        except Exception as exc:
            if PIPE_TABLE_NAME in str(exc) and ("PGRST205" in str(exc) or "schema cache" in str(exc)):
                return pd.DataFrame()
            raise
        df = pd.DataFrame(rows)
    else:
        should_close = conn is None
        conn = conn or get_connection()
        init_db(conn)
        df = pd.read_sql_query("SELECT * FROM pipe_repair_details ORDER BY date, project_sheet, pipe_no", conn)
        if should_close:
            conn.close()

    return _normalize_pipe_repair_details(df)


def load_pipe_repair_details_for_date(selected_date, conn: sqlite3.Connection | None = None) -> pd.DataFrame:
    date_text = pd.to_datetime(selected_date).date().isoformat()
    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        try:
            rows: list[dict[str, Any]] = []
            page_size = 1000
            for start in range(0, 100_000, page_size):
                response = (
                    client.table(PIPE_TABLE_NAME)
                    .select("*")
                    .eq("date", date_text)
                    .order("project_sheet")
                    .range(start, start + page_size - 1)
                    .execute()
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
        "SELECT * FROM pipe_repair_details WHERE date = ? ORDER BY project_sheet, pipe_no",
        conn,
        params=(date_text,),
    )
    if should_close:
        conn.close()
    return _normalize_pipe_repair_details(df)


def _normalize_pipe_repair_details(df: pd.DataFrame) -> pd.DataFrame:
    if not df.empty:
        df["date"] = pd.to_datetime(df["date"])
        placeholder_mask = pd.to_numeric(df["repair_amount"], errors="coerce").le(0.0001)
        df.loc[placeholder_mask, ["repair_amount", "repair_ratio"]] = 0.0
        sort_columns = [column for column in ["date", "project_sheet", "pipe_no"] if column in df.columns]
        if sort_columns:
            df = df.sort_values(sort_columns).reset_index(drop=True)
    return df


def upsert_historical_baselines(df: pd.DataFrame, conn: sqlite3.Connection | None = None) -> int:
    records = _baseline_records_from_df(df)

    if get_backend_name() == "supabase":
        client = _get_supabase_client()
        client.table(BASELINE_TABLE_NAME).upsert(records, on_conflict="project_no,dimensions").execute()
        return len(records)

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    conn.executemany(BASELINE_UPSERT_SQL, records)
    conn.commit()
    if should_close:
        conn.close()
    return len(records)


def load_historical_baselines(conn: sqlite3.Connection | None = None) -> pd.DataFrame:
    if get_backend_name() == "supabase":
        try:
            rows = _fetch_supabase_table(BASELINE_TABLE_NAME, ("project_no", "dimensions"))
        except Exception as exc:
            if "historical_baselines" in str(exc) and ("PGRST205" in str(exc) or "schema cache" in str(exc)):
                return pd.DataFrame()
            raise
        df = pd.DataFrame(rows)
        from baseline import enrich_historical_baseline_metadata

        return enrich_historical_baseline_metadata(df)

    should_close = conn is None
    conn = conn or get_connection()
    init_db(conn)
    df = pd.read_sql_query("SELECT * FROM historical_baselines ORDER BY project_no, dimensions", conn)
    if should_close:
        conn.close()
    from baseline import enrich_historical_baseline_metadata

    return enrich_historical_baseline_metadata(df)
