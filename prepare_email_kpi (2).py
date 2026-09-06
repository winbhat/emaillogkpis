#!/usr/bin/env python3
"""
Prepare Outlook email logs for Average Handling Time (AHT) reporting in Power BI.

AHT definition used:
    Handling Time per conversation =
        timestamp of last considered email in the thread
        - timestamp of first considered email in the thread

    Average Handling Time =
        average of conversation-level handling times

Thread key:
    Unique Subject ID

Team ownership:
    Responsible Team

Period attribution:
    Based on the first considered email timestamp ("Received Timestamp").

Output:
    <output_dir>/powerbi_email_kpi.xlsx
    <output_dir>/aht_thread_fact.csv
    <output_dir>/calendar.csv
    <output_dir>/email_log_clean.csv
    <output_dir>/data_quality.csv

Install:
    pip install pandas openpyxl

Example:
    python prepare_email_kpi.py --input outlook_email_log.xlsx --sheet Sheet1 --output-dir output
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Optional

import pandas as pd


# ---------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------

# Set True if a conversation containing only one considered email should
# NOT participate in Average Handling Time.
EXCLUDE_SINGLE_EMAIL_THREADS = False

# Set True to keep only rows where "To Be Considered" evaluates to True/Yes/1.
FILTER_TO_BE_CONSIDERED = True

# Column aliases make the program tolerant of spaces, underscores,
# capitalization, and common naming variations.
COLUMN_ALIASES = {
    "timestamp": [
        "timestamp", "time stamp", "date time", "datetime", "date_time"
    ],
    "sender": [
        "sender", "from", "from email", "sender email"
    ],
    "to": [
        "to", "recipients", "recipient", "to recipients"
    ],
    "cc": [
        "cc", "cc count", "number of cc", "cc recipients"
    ],
    "message_subject": [
        "message subject", "subject", "email subject"
    ],
    "total_size_kb": [
        "total size", "total size kb", "total size kbs", "size kb",
        "email size kb", "message size kb"
    ],
    "to_be_considered": [
        "to be considered", "considered", "include", "include in kpi",
        "kpi considered"
    ],
    "sender_type": [
        "sender type", "type of sender"
    ],
    "sender_team": [
        "sender team", "team of sender"
    ],
    "recipient_team": [
        "recipient team", "team of recipient"
    ],
    "responsible_team": [
        "responsible team", "owner team", "handling team"
    ],
    "unique_subject_id": [
        "unique subject id", "unique_subject_id", "subject id",
        "conversation id", "thread id"
    ],
    "sequence_number": [
        "sequence number", "sequence", "sequence no", "seq",
        "sequence_number"
    ],
}


# ---------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------

def normalize_name(value: str) -> str:
    """Normalize a column name for matching."""
    value = str(value).strip().lower()
    value = re.sub(r"[_\-]+", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value


def resolve_columns(df: pd.DataFrame) -> dict[str, str]:
    """
    Resolve source Excel columns to canonical names using COLUMN_ALIASES.
    Raises a useful error if mandatory fields are missing.
    """
    normalized_source = {normalize_name(c): c for c in df.columns}
    resolved: dict[str, str] = {}

    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            key = normalize_name(alias)
            if key in normalized_source:
                resolved[canonical] = normalized_source[key]
                break

    mandatory = [
        "timestamp",
        "responsible_team",
        "unique_subject_id",
    ]
    if FILTER_TO_BE_CONSIDERED:
        mandatory.append("to_be_considered")

    missing = [c for c in mandatory if c not in resolved]
    if missing:
        source_cols = ", ".join(map(str, df.columns))
        raise ValueError(
            "Could not identify mandatory column(s): "
            + ", ".join(missing)
            + "\n\nColumns found in the input file:\n"
            + source_cols
            + "\n\nEdit COLUMN_ALIASES near the top of the script if your "
              "headers use different names."
        )

    return resolved


def truthy(value) -> bool:
    """Interpret common Excel representations of Yes/True/1 as True."""
    if pd.isna(value):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value != 0

    text = str(value).strip().lower()
    return text in {
        "true", "yes", "y", "1", "include", "included", "consider",
        "considered", "x"
    }


def clean_team(series: pd.Series) -> pd.Series:
    """Normalize blank team names to 'Unassigned'."""
    s = series.astype("string").str.strip()
    return s.mask(s.isna() | (s == ""), "Unassigned")


def safe_first(series: pd.Series):
    """First non-null value; otherwise NA."""
    non_null = series.dropna()
    return non_null.iloc[0] if len(non_null) else pd.NA


def safe_last(series: pd.Series):
    """Last non-null value; otherwise NA."""
    non_null = series.dropna()
    return non_null.iloc[-1] if len(non_null) else pd.NA


def make_calendar(min_date: pd.Timestamp, max_date: pd.Timestamp) -> pd.DataFrame:
    """Create a Power BI-friendly date dimension."""
    dates = pd.date_range(min_date.normalize(), max_date.normalize(), freq="D")
    cal = pd.DataFrame({"Date": dates})

    iso = cal["Date"].dt.isocalendar()

    cal["Year"] = cal["Date"].dt.year
    cal["Quarter Number"] = cal["Date"].dt.quarter
    cal["Quarter"] = (
        cal["Year"].astype(str)
        + "-Q"
        + cal["Quarter Number"].astype(str)
    )
    cal["Month Number"] = cal["Date"].dt.month
    cal["Month"] = cal["Date"].dt.strftime("%Y-%m")
    cal["Month Name"] = cal["Date"].dt.strftime("%B")
    cal["Month Start"] = cal["Date"].dt.to_period("M").dt.start_time

    cal["ISO Year"] = iso["year"].astype(int)
    cal["ISO Week"] = iso["week"].astype(int)
    cal["Year Week"] = (
        cal["ISO Year"].astype(str)
        + "-W"
        + cal["ISO Week"].astype(str).str.zfill(2)
    )
    cal["Week Start"] = cal["Date"] - pd.to_timedelta(
        cal["Date"].dt.weekday, unit="D"
    )

    cal["Day"] = cal["Date"].dt.day
    cal["Day Name"] = cal["Date"].dt.strftime("%A")
    cal["Day of Week Number"] = cal["Date"].dt.weekday + 1
    cal["Is Weekend"] = cal["Date"].dt.weekday >= 5

    return cal


def weekday_elapsed_hours(start: pd.Timestamp, end: pd.Timestamp) -> float:
    """Return elapsed hours excluding all time on Saturday and Sunday."""
    if pd.isna(start) or pd.isna(end) or end < start:
        return float("nan")
    if end == start:
        return 0.0

    total_seconds = 0.0
    cursor = start
    while cursor < end:
        next_midnight = cursor.normalize() + pd.Timedelta(days=1)
        segment_end = min(next_midnight, end)
        if cursor.weekday() < 5:
            total_seconds += (segment_end - cursor).total_seconds()
        cursor = segment_end
    return total_seconds / 3600.0


# ---------------------------------------------------------------------
# CORE TRANSFORMATION
# ---------------------------------------------------------------------

def prepare_data(input_path: Path, sheet_name: Optional[str]) -> tuple[
    pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame
]:
    """Read input Excel and prepare clean log, thread fact, calendar, quality."""

    # Read Excel
    if sheet_name:
        raw = pd.read_excel(input_path, sheet_name=sheet_name)
    else:
        raw = pd.read_excel(input_path, sheet_name=0)

    raw.columns = [str(c).strip() for c in raw.columns]
    resolved = resolve_columns(raw)

    # Rename matched fields to canonical internal names.
    rename_map = {source: canonical for canonical, source in resolved.items()}
    df = raw.rename(columns=rename_map).copy()

    # Guarantee optional columns exist so downstream logic remains stable.
    optional_fields = [
        "sender", "to", "cc", "message_subject", "total_size_kb",
        "sender_type", "sender_team", "recipient_team", "sequence_number"
    ]
    for col in optional_fields:
        if col not in df.columns:
            df[col] = pd.NA

    # Parse timestamp.
    df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce")

    # Normalize thread IDs and team.
    df["unique_subject_id"] = (
        df["unique_subject_id"].astype("string").str.strip()
    )
    df["responsible_team"] = clean_team(df["responsible_team"])

    # KPI inclusion flag.
    if "to_be_considered" in df.columns:
        df["is_considered"] = df["to_be_considered"].map(truthy)
    else:
        df["is_considered"] = True

    # Base data quality flags before filtering.
    df["dq_invalid_timestamp"] = df["timestamp"].isna()
    df["dq_missing_thread_id"] = (
        df["unique_subject_id"].isna()
        | (df["unique_subject_id"] == "")
    )
    df["dq_unassigned_team"] = df["responsible_team"].eq("Unassigned")

    # Keep only rows eligible for KPI if requested.
    work = df.copy()
    if FILTER_TO_BE_CONSIDERED:
        work = work.loc[work["is_considered"]].copy()

    # Rows without timestamp/thread ID cannot be used for AHT.
    valid = work.loc[
        work["timestamp"].notna()
        & work["unique_subject_id"].notna()
        & work["unique_subject_id"].ne("")
    ].copy()

    # Timestamp controls sequence; input sequence number is retained only
    # for comparison/debugging.
    valid = valid.sort_values(
        ["unique_subject_id", "timestamp"],
        kind="stable"
    )

    # Recalculate deterministic sequence from timestamp.
    valid["calculated_sequence_number"] = (
        valid.groupby("unique_subject_id").cumcount() + 1
    )

    # First response = timestamp of sequence number 2 in timestamp order.
    first_response = (
        valid.loc[
            valid["calculated_sequence_number"].eq(2),
            ["unique_subject_id", "timestamp"]
        ]
        .rename(columns={"timestamp": "first_response_timestamp"})
    )

    # Detect threads with more than one responsible team.
    team_counts = (
        valid.groupby("unique_subject_id")["responsible_team"]
        .nunique(dropna=True)
        .rename("distinct_responsible_team_count")
    )

    # One row per conversation/thread.
    thread = (
        valid.groupby("unique_subject_id", as_index=False)
        .agg(
            received_timestamp=("timestamp", "min"),
            resolution_timestamp=("timestamp", "max"),
            email_count=("timestamp", "size"),
            responsible_team=("responsible_team", safe_first),
            first_sender=("sender", safe_first),
            last_sender=("sender", safe_last),
            message_subject=("message_subject", safe_first),
            first_sender_type=("sender_type", safe_first),
            first_sender_team=("sender_team", safe_first),
            first_recipient_team=("recipient_team", safe_first),
            total_thread_size_kb=("total_size_kb", "sum"),
        )
    )

    thread = thread.merge(
        team_counts.reset_index(),
        on="unique_subject_id",
        how="left"
    )

    thread = thread.merge(
        first_response,
        on="unique_subject_id",
        how="left"
    )

    # First-response metrics.
    thread["first_response_hours"] = (
        thread["first_response_timestamp"] - thread["received_timestamp"]
    ).dt.total_seconds() / 3600.0
    thread["first_response_days"] = thread["first_response_hours"] / 24.0

    thread["first_response_weekday_hours"] = thread.apply(
        lambda row: weekday_elapsed_hours(
            row["received_timestamp"],
            row["first_response_timestamp"]
        ),
        axis=1
    )
    thread["first_response_weekday_days"] = (
        thread["first_response_weekday_hours"] / 24.0
    )

    # Nullable Boolean: blank when there is no second email.
    thread["first_response_within_48h_sla"] = pd.Series(
        pd.NA, index=thread.index, dtype="boolean"
    )
    has_response = thread["first_response_timestamp"].notna()
    thread.loc[has_response, "first_response_within_48h_sla"] = (
        thread.loc[has_response, "first_response_weekday_hours"] <= 48.0
    )

    # AHT duration.
    thread["handling_seconds"] = (
        thread["resolution_timestamp"] - thread["received_timestamp"]
    ).dt.total_seconds()

    thread["handling_minutes"] = thread["handling_seconds"] / 60.0
    thread["handling_hours"] = thread["handling_seconds"] / 3600.0
    thread["handling_days"] = thread["handling_seconds"] / 86400.0

    thread["is_single_email_thread"] = thread["email_count"].eq(1)
    thread["has_multiple_responsible_teams"] = (
        thread["distinct_responsible_team_count"] > 1
    )

    # Optional exclusion of one-message threads.
    if EXCLUDE_SINGLE_EMAIL_THREADS:
        thread["include_in_aht"] = ~thread["is_single_email_thread"]
    else:
        thread["include_in_aht"] = True

    # Date dimensions are attributed to the FIRST email / received timestamp.
    thread["received_date"] = thread["received_timestamp"].dt.normalize()
    thread["resolution_date"] = thread["resolution_timestamp"].dt.normalize()

    iso = thread["received_timestamp"].dt.isocalendar()

    thread["received_year"] = thread["received_timestamp"].dt.year
    thread["received_quarter_number"] = (
        thread["received_timestamp"].dt.quarter
    )
    thread["received_quarter"] = (
        thread["received_year"].astype("Int64").astype(str)
        + "-Q"
        + thread["received_quarter_number"].astype("Int64").astype(str)
    )
    thread["received_month_number"] = (
        thread["received_timestamp"].dt.month
    )
    thread["received_month"] = thread["received_timestamp"].dt.strftime(
        "%Y-%m"
    )
    thread["received_month_start"] = (
        thread["received_timestamp"].dt.to_period("M").dt.start_time
    )
    thread["received_iso_year"] = iso["year"].astype("Int64")
    thread["received_iso_week"] = iso["week"].astype("Int64")
    thread["received_year_week"] = (
        thread["received_iso_year"].astype(str)
        + "-W"
        + thread["received_iso_week"].astype(str).str.zfill(2)
    )
    thread["received_week_start"] = (
        thread["received_date"]
        - pd.to_timedelta(thread["received_date"].dt.weekday, unit="D")
    )

    # Sort the fact table in a predictable order.
    thread = thread.sort_values(
        ["received_timestamp", "unique_subject_id"]
    ).reset_index(drop=True)

    # Power BI calendar.
    if len(thread):
        min_date = min(
            thread["received_date"].min(),
            thread["resolution_date"].min()
        )
        max_date = max(
            thread["received_date"].max(),
            thread["resolution_date"].max()
        )
        calendar = make_calendar(min_date, max_date)
    else:
        calendar = pd.DataFrame(columns=["Date"])

    # Data quality summary.
    quality_rows = [
        ("Input rows", len(df)),
        ("Rows marked considered", int(df["is_considered"].sum())),
        ("Rows used after considered filter", len(work)),
        ("Rows with invalid timestamp", int(df["dq_invalid_timestamp"].sum())),
        ("Rows with missing thread ID", int(df["dq_missing_thread_id"].sum())),
        ("Rows with unassigned responsible team", int(df["dq_unassigned_team"].sum())),
        ("Valid rows used for AHT", len(valid)),
        ("Unique conversations", len(thread)),
        ("Single-email conversations", int(thread["is_single_email_thread"].sum()) if len(thread) else 0),
        ("Threads with multiple responsible teams", int(thread["has_multiple_responsible_teams"].sum()) if len(thread) else 0),
    ]
    quality = pd.DataFrame(quality_rows, columns=["Metric", "Value"])

    # Friendly column names for Power BI.
    thread = thread.rename(columns={
        "unique_subject_id": "Unique Subject ID",
        "received_timestamp": "Received Timestamp",
        "resolution_timestamp": "Resolution Timestamp",
        "email_count": "Email Count",
        "responsible_team": "Responsible Team",
        "first_sender": "First Sender",
        "last_sender": "Last Sender",
        "message_subject": "Message Subject",
        "first_sender_type": "First Sender Type",
        "first_sender_team": "First Sender Team",
        "first_recipient_team": "First Recipient Team",
        "total_thread_size_kb": "Total Thread Size KB",
        "distinct_responsible_team_count": "Distinct Responsible Team Count",
        "first_response_timestamp": "First Response Timestamp",
        "first_response_hours": "First Response Hours",
        "first_response_days": "First Response Days",
        "first_response_weekday_hours": "First Response Weekday Hours",
        "first_response_weekday_days": "First Response Weekday Days",
        "first_response_within_48h_sla": "First Response Within 48h SLA",
        "handling_seconds": "Handling Seconds",
        "handling_minutes": "Handling Minutes",
        "handling_hours": "Handling Hours",
        "handling_days": "Handling Days",
        "is_single_email_thread": "Is Single Email Thread",
        "has_multiple_responsible_teams": "Has Multiple Responsible Teams",
        "include_in_aht": "Include In AHT",
        "received_date": "Received Date",
        "resolution_date": "Resolution Date",
        "received_year": "Received Year",
        "received_quarter_number": "Received Quarter Number",
        "received_quarter": "Received Quarter",
        "received_month_number": "Received Month Number",
        "received_month": "Received Month",
        "received_month_start": "Received Month Start",
        "received_iso_year": "Received ISO Year",
        "received_iso_week": "Received ISO Week",
        "received_year_week": "Received Year Week",
        "received_week_start": "Received Week Start",
    })

    # Clean log output with source detail plus recalculated sequence.
    clean_log = valid.rename(columns={
        "timestamp": "Timestamp",
        "sender": "Sender",
        "to": "To",
        "cc": "CC",
        "message_subject": "Message Subject",
        "total_size_kb": "Total Size KB",
        "to_be_considered": "To Be Considered",
        "sender_type": "Sender Type",
        "sender_team": "Sender Team",
        "recipient_team": "Recipient Team",
        "responsible_team": "Responsible Team",
        "unique_subject_id": "Unique Subject ID",
        "sequence_number": "Source Sequence Number",
        "calculated_sequence_number": "Calculated Sequence Number",
    })

    return clean_log, thread, calendar, quality


# ---------------------------------------------------------------------
# OUTPUT
# ---------------------------------------------------------------------

def write_outputs(
    clean_log: pd.DataFrame,
    thread: pd.DataFrame,
    calendar: pd.DataFrame,
    quality: pd.DataFrame,
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    xlsx_path = output_dir / "powerbi_email_kpi.xlsx"

    # Excel is convenient for direct Power BI import / inspection.
    with pd.ExcelWriter(
        xlsx_path,
        engine="openpyxl",
        datetime_format="yyyy-mm-dd hh:mm:ss"
    ) as writer:
        thread.to_excel(writer, sheet_name="AHT_Thread_Fact", index=False)
        calendar.to_excel(writer, sheet_name="Calendar", index=False)
        clean_log.to_excel(writer, sheet_name="Email_Log_Clean", index=False)
        quality.to_excel(writer, sheet_name="Data_Quality", index=False)

        # Lightweight usability formatting.
        for sheet_name in [
            "AHT_Thread_Fact", "Calendar", "Email_Log_Clean", "Data_Quality"
        ]:
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            # Sensible widths without expensive full-cell autosizing.
            for col_cells in ws.iter_cols(
                min_row=1,
                max_row=min(ws.max_row, 200),
                min_col=1,
                max_col=ws.max_column
            ):
                max_len = 0
                for cell in col_cells:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                width = min(max(max_len + 2, 11), 35)
                ws.column_dimensions[col_cells[0].column_letter].width = width

    # CSV copies are often faster for large Power BI models.
    thread.to_csv(output_dir / "aht_thread_fact.csv", index=False)
    calendar.to_csv(output_dir / "calendar.csv", index=False)
    clean_log.to_csv(output_dir / "email_log_clean.csv", index=False)
    quality.to_csv(output_dir / "data_quality.csv", index=False)

    print("\nCreated:")
    print(f"  {xlsx_path}")
    print(f"  {output_dir / 'aht_thread_fact.csv'}")
    print(f"  {output_dir / 'calendar.csv'}")
    print(f"  {output_dir / 'email_log_clean.csv'}")
    print(f"  {output_dir / 'data_quality.csv'}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Prepare Outlook email logs for Power BI AHT reporting."
    )
    parser.add_argument(
        "--input",
        required=True,
        help="Path to the source .xlsx file."
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Excel sheet name. Defaults to the first sheet."
    )
    parser.add_argument(
        "--output-dir",
        default="powerbi_output",
        help="Directory for generated files. Default: powerbi_output"
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    clean_log, thread, calendar, quality = prepare_data(
        input_path=input_path,
        sheet_name=args.sheet
    )

    write_outputs(
        clean_log=clean_log,
        thread=thread,
        calendar=calendar,
        quality=quality,
        output_dir=output_dir
    )

    eligible = thread.loc[thread["Include In AHT"]]
    if len(eligible):
        overall_aht_hours = eligible["Handling Hours"].mean()
        print(
            f"\nOverall Average Handling Time: "
            f"{overall_aht_hours:.2f} hours"
        )
    else:
        print("\nNo eligible conversations were available for AHT.")


if __name__ == "__main__":
    main()
