#!/usr/bin/env python3
"""
Populate the Molt Canon workbook Raw Import sheet from the live canon API.

Usage:
    python populate_molt_workbook.py \
        --workbook molt_extraction_workbook_and_shortlist_model_full_impl_seed71.xlsx \
        --output molt_extraction_workbook_and_shortlist_model_live.xlsx \
        --snapshot Snapshot_002

This script only writes the Raw Import sheet and import metadata.
Editorial screening, scoring, shortlist review, and anthology curation
should still be done in the workbook.
"""
import argparse
import datetime as dt
import json
import requests
from openpyxl import load_workbook

API_URL = "https://molt.church/api/canon"

def fetch_canon():
    resp = requests.get(API_URL, timeout=30)
    resp.raise_for_status()
    payload = resp.json()
    return payload.get("the_great_book", [])

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--snapshot", required=True)
    args = ap.parse_args()

    records = fetch_canon()
    wb = load_workbook(args.workbook)
    ws = wb["Raw Import"]

    # clear prior import rows
    for r in range(5, ws.max_row + 1):
        for c in range(1, 10):
            ws.cell(r, c).value = None

    imported_at = dt.datetime.utcnow().replace(microsecond=0).isoformat()
    batch_name = f"live_full_import_{imported_at}"

    for i, item in enumerate(records, start=5):
        ws.cell(i, 1).value = i - 4
        ws.cell(i, 2).value = item.get("prophet_name")
        ws.cell(i, 3).value = item.get("scripture_type")
        ws.cell(i, 4).value = item.get("content")
        ws.cell(i, 5).value = item.get("canonized_at")
        ws.cell(i, 6).value = API_URL
        ws.cell(i, 7).value = args.snapshot
        ws.cell(i, 8).value = imported_at
        ws.cell(i, 9).value = batch_name

    # update light control metadata
    cws = wb["Controls"]
    cws["B5"] = API_URL
    cws["B6"] = args.snapshot
    cws["A8"] = "Imported live row count"
    cws["B8"] = len(records)
    cws["C8"] = "Updated by import script"

    wb.save(args.output)

if __name__ == "__main__":
    main()
