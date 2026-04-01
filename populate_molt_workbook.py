#!/usr/bin/env python3
"""Production loader for Molt canon workbook full imports."""

import argparse
import datetime as dt
import json
from collections import Counter
from copy import copy
from pathlib import Path
from typing import Any

import requests
from openpyxl import load_workbook

API_URL = "https://molt.church/api/canon"
RAW_IMPORT_SHEET = "Raw Import"
CONTROL_SHEET = "Controls"
START_ROW = 5
RAW_COLUMNS = 9


class CanonLoadError(RuntimeError):
    """Raised when a live canon load cannot be completed."""


def utc_now() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0)


def to_excel_cell(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    return str(value)


def normalize_canonized_at(value: Any, anomalies: list[str], row_id: int) -> str:
    raw = to_excel_cell(value)
    if not raw:
        anomalies.append(f"row {row_id}: missing canonized_at")
        return ""

    # Accept unmodified source values but flag malformed dates.
    try:
        dt.datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except ValueError:
        anomalies.append(f"row {row_id}: malformed canonized_at '{raw}'")
    return raw


def fetch_payload(source_url: str, timeout_seconds: int) -> dict[str, Any]:
    response = requests.get(source_url, timeout=timeout_seconds)
    response.raise_for_status()
    payload = response.json()
    if not isinstance(payload, dict):
        raise CanonLoadError("API payload root is not a JSON object")
    return payload


def flatten_records(payload: dict[str, Any]) -> tuple[list[dict[str, Any]], list[str], dict[str, int]]:
    anomalies: list[str] = []
    book = payload.get("the_great_book")
    if book is None:
        raise CanonLoadError("Payload missing required 'the_great_book' collection")
    if not isinstance(book, list):
        raise CanonLoadError("Payload field 'the_great_book' is not a list")

    rows: list[dict[str, Any]] = []
    scripture_types = Counter()
    content_seen: Counter[str] = Counter()

    for idx, item in enumerate(book, start=1):
        if not isinstance(item, dict):
            anomalies.append(f"row {idx}: non-object entry skipped (type={type(item).__name__})")
            continue

        prophet_name = to_excel_cell(item.get("prophet_name"))
        scripture_type = to_excel_cell(item.get("scripture_type"))
        content = to_excel_cell(item.get("content"))
        canonized_at = normalize_canonized_at(item.get("canonized_at"), anomalies, idx)

        if not prophet_name:
            anomalies.append(f"row {idx}: missing prophet_name")
        if not scripture_type:
            anomalies.append(f"row {idx}: missing scripture_type")
            scripture_types["<missing>"] += 1
        else:
            scripture_types[scripture_type] += 1
        if not content:
            anomalies.append(f"row {idx}: missing content")

        content_seen[content] += 1
        rows.append(
            {
                "Raw_Row_ID": len(rows) + 1,
                "Raw_Prophet_Name": prophet_name,
                "Raw_Scripture_Type": scripture_type,
                "Raw_Content": content,
                "Raw_Canonized_At": canonized_at,
            }
        )

    duplicate_non_empty = sum(1 for k, count in content_seen.items() if k and count > 1)
    if duplicate_non_empty:
        anomalies.append(f"detected {duplicate_non_empty} duplicate Raw_Content values")

    stats = {
        "source_entries": len(book),
        "imported_entries": len(rows),
        "skipped_entries": len(book) - len(rows),
    }
    return rows, anomalies, dict(scripture_types)


def clear_raw_import(ws, keep_template_row: int = START_ROW) -> None:
    max_row = ws.max_row
    for row in range(keep_template_row, max_row + 1):
        for col in range(1, RAW_COLUMNS + 1):
            cell = ws.cell(row=row, column=col)
            if row == keep_template_row:
                # Retain style template on first data row; reset value only.
                cell.value = None
            else:
                ws.cell(row=row, column=col).value = None


def copy_row_style(ws, from_row: int, to_row: int) -> None:
    for col in range(1, RAW_COLUMNS + 1):
        src = ws.cell(row=from_row, column=col)
        dst = ws.cell(row=to_row, column=col)
        dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)


def write_rows(
    ws,
    rows: list[dict[str, Any]],
    source_url: str,
    snapshot_id: str,
    imported_at: str,
    import_batch_name: str,
) -> None:
    clear_raw_import(ws)

    for offset, item in enumerate(rows):
        row_num = START_ROW + offset
        if row_num > START_ROW:
            copy_row_style(ws, START_ROW, row_num)

        ws.cell(row=row_num, column=1, value=item["Raw_Row_ID"])
        ws.cell(row=row_num, column=2, value=item["Raw_Prophet_Name"])
        ws.cell(row=row_num, column=3, value=item["Raw_Scripture_Type"])
        ws.cell(row=row_num, column=4, value=item["Raw_Content"])
        ws.cell(row=row_num, column=5, value=item["Raw_Canonized_At"])
        ws.cell(row=row_num, column=6, value=source_url)
        ws.cell(row=row_num, column=7, value=snapshot_id)
        ws.cell(row=row_num, column=8, value=imported_at)
        ws.cell(row=row_num, column=9, value=import_batch_name)


def update_controls(wb, source_url: str, snapshot_id: str, imported_count: int) -> None:
    ws = wb[CONTROL_SHEET]
    ws["B5"] = source_url
    ws["B6"] = snapshot_id
    ws["A8"] = "Imported live row count"
    ws["B8"] = imported_count
    ws["C8"] = "Updated by populate_molt_workbook.py"


def append_run_log(
    path: Path,
    status: str,
    workbook_in: str,
    workbook_out: str,
    snapshot_name: str,
    snapshot_id: str,
    stats: dict[str, int] | None,
    anomalies: list[str],
    assumptions: list[str],
) -> None:
    stamp = utc_now().isoformat()
    lines = [
        "\n## Run - " + stamp,
        f"**Status:** {status}",
        "",
        "### Inputs",
        f"- workbook: {workbook_in}",
        "- script: populate_molt_workbook.py",
        f"- source URL: {API_URL}",
        "",
        "### Outputs",
        f"- workbook: {workbook_out}",
        f"- snapshot: {snapshot_name}",
        f"- snapshot ID: {snapshot_id}",
        "",
        "### Counts",
    ]
    if stats:
        lines.extend(
            [
                f"- rows fetched: {stats['source_entries']}",
                f"- rows imported: {stats['imported_entries']}",
                f"- rows skipped: {stats['skipped_entries']}",
            ]
        )
    else:
        lines.append("- unavailable due to failed fetch")

    lines.extend(["", "### Schema anomalies"]) 
    if anomalies:
        lines.extend([f"- {entry}" for entry in anomalies])
    else:
        lines.append("- none")

    lines.extend(["", "### Assumptions"]) 
    if assumptions:
        lines.extend([f"- {entry}" for entry in assumptions])
    else:
        lines.append("- none")

    path.write_text(path.read_text() + "\n".join(lines) + "\n", encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Load full Molt canon into workbook Raw Import")
    parser.add_argument("--workbook", required=True, help="Input workbook path")
    parser.add_argument("--output", default="molt_extraction_workbook_and_shortlist_model_full_canon.xlsx")
    parser.add_argument("--source-url", default=API_URL)
    parser.add_argument("--timeout", type=int, default=30)
    parser.add_argument("--run-log", default="RUN_LOG.md")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    now = utc_now()
    snapshot_id = f"Snapshot_{now.strftime('%Y%m%d_%H%M%S')}"
    snapshot_name = f"molt_canon_snapshot_{now.strftime('%Y%m%d_%H%M%S')}.json"
    imported_at = now.isoformat()
    batch_name = f"live_full_canon_{now.strftime('%Y%m%d_%H%M%S')}"

    assumptions = [
        "the_great_book entries are canonical source rows in order provided by API",
        "non-object entries are invalid and skipped with explicit logging",
    ]

    try:
        payload = fetch_payload(args.source_url, args.timeout)
    except Exception as exc:
        append_run_log(
            Path(args.run_log),
            status=f"FAILED - fetch error: {exc}",
            workbook_in=args.workbook,
            workbook_out=args.output,
            snapshot_name="not_written",
            snapshot_id=snapshot_id,
            stats=None,
            anomalies=[f"fetch_error: {exc}"],
            assumptions=assumptions,
        )
        raise

    Path(snapshot_name).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    rows, anomalies, scripture_types = flatten_records(payload)
    anomalies.extend([f"scripture_type '{k}': {v} rows" for k, v in sorted(scripture_types.items())])

    wb = load_workbook(args.workbook)
    raw_ws = wb[RAW_IMPORT_SHEET]
    write_rows(raw_ws, rows, args.source_url, snapshot_id, imported_at, batch_name)
    update_controls(wb, args.source_url, snapshot_id, len(rows))
    wb.save(args.output)

    stats = {
        "source_entries": len(payload.get("the_great_book", [])),
        "imported_entries": len(rows),
        "skipped_entries": len(payload.get("the_great_book", [])) - len(rows),
    }

    append_run_log(
        Path(args.run_log),
        status="SUCCESS",
        workbook_in=args.workbook,
        workbook_out=args.output,
        snapshot_name=snapshot_name,
        snapshot_id=snapshot_id,
        stats=stats,
        anomalies=anomalies,
        assumptions=assumptions,
    )


if __name__ == "__main__":
    main()
