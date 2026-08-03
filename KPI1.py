import re
from pathlib import Path

import pandas as pd


# ============================================================
# CONFIGURATION
# ============================================================

INPUT_FILE = Path("email_log.xlsx")
OUTPUT_FILE = Path("email_kpi_results.xlsx")
SHEET_NAME = 0  # First sheet, or replace with a sheet name

# Change these names to match your Excel headers exactly.
COLUMN_MAP = {
    "Timestamp": "timestamp",
    "sender": "sender",
    "TO": "recipient",
    "message subject": "subject",
    "sender type": "sender_type",
    "sender team": "sender_team",
    "receiver team": "receiver_team",
    "responsible team": "responsible_team",
}


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def normalize_subject(subject: object) -> str:
    """
    Normalize an email subject so replies and forwards are grouped together.

    Examples:
        'RE: Invoice problem'       -> 'invoice problem'
        'FW: RE: Invoice problem'   -> 'invoice problem'
    """
    if pd.isna(subject):
        return ""

    subject = str(subject).strip().lower()

    # Repeatedly remove prefixes such as RE:, FW:, FWD:
    prefix_pattern = r"^\s*((re|fw|fwd)\s*:\s*)+"
    subject = re.sub(prefix_pattern, "", subject, flags=re.IGNORECASE)

    # Remove common external-email tags.
    subject = re.sub(
        r"^\s*\[(external|ext)\]\s*",
        "",
        subject,
        flags=re.IGNORECASE,
    )

    # Replace multiple spaces with one.
    subject = re.sub(r"\s+", " ", subject).strip()

    return subject


def first_non_empty(series: pd.Series):
    """Return the first non-empty value in a pandas Series."""
    values = series.dropna().astype(str).str.strip()
    values = values[values.ne("")]
    return values.iloc[0] if not values.empty else pd.NA


def calculate_email_kpis(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Calculate handling-time KPIs for each email conversation.

    Main assumption:
    - First email in the conversation = receipt/start.
    - Last email in the conversation = resolution/end.
    - Handling time = last timestamp - first timestamp.
    """

    required_columns = [
        "timestamp",
        "sender",
        "recipient",
        "subject",
        "sender_type",
        "sender_team",
        "receiver_team",
        "responsible_team",
    ]

    missing_columns = [
        column for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            f"Missing required columns after renaming: {missing_columns}"
        )

    data = df.copy()

    # Convert timestamps to real datetime values.
    data["timestamp"] = pd.to_datetime(
        data["timestamp"],
        errors="coerce",
        dayfirst=True,  # Change to False if dates are month/day/year.
    )

    # Remove rows without a valid timestamp.
    invalid_timestamp_count = data["timestamp"].isna().sum()

    if invalid_timestamp_count:
        print(
            f"Warning: {invalid_timestamp_count} rows were removed "
            "because their timestamps could not be parsed."
        )

    data = data.dropna(subset=["timestamp"]).copy()

    # Standardize sender type.
    data["sender_type"] = (
        data["sender_type"]
        .astype("string")
        .str.strip()
        .str.lower()
    )

    # Normalize the subject to create an estimated conversation ID.
    data["normalized_subject"] = data["subject"].apply(normalize_subject)

    # Avoid grouping all blank subjects into one conversation.
    blank_subject = data["normalized_subject"].eq("")

    data.loc[blank_subject, "normalized_subject"] = (
        "blank-subject-row-"
        + data.index[blank_subject].astype(str)
    )

    # Estimated conversation identifier.
    # Prefer an actual Thread ID/Conversation ID when available.
    data["conversation_id"] = data["normalized_subject"]

    # Sort chronologically inside every conversation.
    data = data.sort_values(
        ["conversation_id", "timestamp"],
        kind="stable",
    ).reset_index(drop=True)

    conversation_results = []

    for conversation_id, group in data.groupby(
        "conversation_id",
        sort=False,
    ):
        group = group.sort_values("timestamp", kind="stable").copy()

        first_email = group.iloc[0]
        last_email = group.iloc[-1]

        start_time = first_email["timestamp"]
        resolution_time = last_email["timestamp"]
        handling_time = resolution_time - start_time

        # First internal email after the first external email.
        external_emails = group[group["sender_type"].eq("external")]

        first_external_time = (
            external_emails["timestamp"].iloc[0]
            if not external_emails.empty
            else pd.NaT
        )

        first_internal_response_time = pd.NaT

        if pd.notna(first_external_time):
            internal_responses = group[
                group["sender_type"].eq("internal")
                & group["timestamp"].gt(first_external_time)
            ]

            if not internal_responses.empty:
                first_internal_response_time = (
                    internal_responses["timestamp"].iloc[0]
                )

        if (
            pd.notna(first_external_time)
            and pd.notna(first_internal_response_time)
        ):
            first_response_time = (
                first_internal_response_time - first_external_time
            )
        else:
            first_response_time = pd.NaT

        conversation_results.append(
            {
                "conversation_id": conversation_id,
                "original_subject": first_email["subject"],
                "receipt_timestamp": start_time,
                "resolution_timestamp": resolution_time,
                "handling_time": handling_time,
                "handling_time_minutes": handling_time.total_seconds() / 60,
                "handling_time_hours": handling_time.total_seconds() / 3600,
                "handling_time_days": handling_time.total_seconds() / 86400,
                "email_count": len(group),
                "first_sender": first_email["sender"],
                "last_sender": last_email["sender"],
                "last_sender_type": last_email["sender_type"],
                "first_external_timestamp": first_external_time,
                "first_internal_response_timestamp": (
                    first_internal_response_time
                ),
                "first_response_time": first_response_time,
                "first_response_minutes": (
                    first_response_time.total_seconds() / 60
                    if pd.notna(first_response_time)
                    else pd.NA
                ),
                "sender_team": first_non_empty(group["sender_team"]),
                "receiver_team": first_non_empty(group["receiver_team"]),
                "responsible_team": first_non_empty(
                    group["responsible_team"]
                ),
                "considered_resolved": True,
            }
        )

    conversations = pd.DataFrame(conversation_results)

    # Team-level KPI summary.
    team_summary = (
        conversations
        .groupby("responsible_team", dropna=False)
        .agg(
            conversation_count=("conversation_id", "nunique"),
            total_email_count=("email_count", "sum"),
            average_emails_per_conversation=("email_count", "mean"),
            average_handling_minutes=("handling_time_minutes", "mean"),
            median_handling_minutes=("handling_time_minutes", "median"),
            average_handling_hours=("handling_time_hours", "mean"),
            median_handling_hours=("handling_time_hours", "median"),
            maximum_handling_hours=("handling_time_hours", "max"),
            average_first_response_minutes=(
                "first_response_minutes",
                "mean",
            ),
        )
        .reset_index()
    )

    return conversations, team_summary


# ============================================================
# MAIN PROCESS
# ============================================================

def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input file not found: {INPUT_FILE.resolve()}"
        )

    # Read the Excel file.
    raw_data = pd.read_excel(
        INPUT_FILE,
        sheet_name=SHEET_NAME,
    )

    # Clean whitespace from Excel headers.
    raw_data.columns = raw_data.columns.astype(str).str.strip()

    # Rename columns to consistent internal names.
    missing_source_columns = [
        column for column in COLUMN_MAP
        if column not in raw_data.columns
    ]

    if missing_source_columns:
        raise ValueError(
            "These Excel columns were not found: "
            f"{missing_source_columns}\n"
            f"Available columns: {raw_data.columns.tolist()}"
        )

    data = raw_data.rename(columns=COLUMN_MAP)

    conversations, team_summary = calculate_email_kpis(data)

    overall_summary = pd.DataFrame(
        {
            "KPI": [
                "Number of conversations",
                "Number of emails",
                "Average handling time — minutes",
                "Median handling time — minutes",
                "Average handling time — hours",
                "Median handling time — hours",
                "Average first response time — minutes",
            ],
            "Value": [
                conversations["conversation_id"].nunique(),
                conversations["email_count"].sum(),
                conversations["handling_time_minutes"].mean(),
                conversations["handling_time_minutes"].median(),
                conversations["handling_time_hours"].mean(),
                conversations["handling_time_hours"].median(),
                conversations["first_response_minutes"].mean(),
            ],
        }
    )

    # Export the results into a new Excel workbook.
    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl",
        datetime_format="yyyy-mm-dd hh:mm:ss",
    ) as writer:
        conversations.to_excel(
            writer,
            sheet_name="Conversation KPIs",
            index=False,
        )

        team_summary.to_excel(
            writer,
            sheet_name="Team Summary",
            index=False,
        )

        overall_summary.to_excel(
            writer,
            sheet_name="Overall Summary",
            index=False,
        )

    print(f"Results saved to: {OUTPUT_FILE.resolve()}")
    print()
    print(overall_summary.to_string(index=False))


if __name__ == "__main__":
    main()
