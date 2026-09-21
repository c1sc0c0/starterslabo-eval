#!/usr/bin/env python3
"""Copy a month tab in a Starters Labo evaluatiefiche workbook and fill cells from JSON."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment

WRAP_PREFIXES = ("C", "D")


def parse_date(value: str | date | datetime | None) -> date | datetime | str | None:
    if value is None or value == "":
        return date.today()
    if isinstance(value, (date, datetime)):
        return value
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return s


def wrap(ws, coord: str) -> None:
    cell = ws[coord]
    if isinstance(cell.value, str) and len(cell.value) > 40:
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--workbook", required=True, help="Path to evaluatiefiche .xlsx")
    ap.add_argument("--tab", required=True, help="New tab name, e.g. 07_26")
    ap.add_argument(
        "--from-tab",
        help="Source tab to copy (default: last MM_YY sheet in the workbook)",
    )
    ap.add_argument("--answers", required=True, help='JSON file: {"C12": "...", ...}')
    ap.add_argument("--force", action="store_true", help="Overwrite existing --tab")
    args = ap.parse_args()

    path = Path(args.workbook)
    if not path.exists():
        print(f"ERROR: workbook not found: {path}", file=sys.stderr)
        return 1

    answers = json.loads(Path(args.answers).read_text(encoding="utf-8"))
    if not isinstance(answers, dict):
        print("ERROR: answers JSON must be an object of cell→value", file=sys.stderr)
        return 1

    wb = load_workbook(path)

    if args.tab in wb.sheetnames:
        if not args.force:
            print(
                f"ERROR: tab {args.tab!r} already exists (pass --force to replace)",
                file=sys.stderr,
            )
            return 1
        del wb[args.tab]

    src_name = args.from_tab
    if not src_name:
        month_tabs = [
            n
            for n in wb.sheetnames
            if len(n) == 5 and n[2] == "_" and n[:2].isdigit()
        ]
        if not month_tabs:
            print("ERROR: no MM_YY source tab found; pass --from-tab", file=sys.stderr)
            return 1
        src_name = month_tabs[-1]

    if src_name not in wb.sheetnames:
        print(f"ERROR: source tab {src_name!r} missing; have {wb.sheetnames}", file=sys.stderr)
        return 1

    ws = wb.copy_worksheet(wb[src_name])
    ws.title = args.tab

    target_idx = wb.sheetnames.index(src_name) + 1
    current_idx = wb.sheetnames.index(args.tab)
    wb.move_sheet(ws, offset=target_idx - current_idx)

    for coord, value in answers.items():
        coord = str(coord).upper()
        if coord == "B7":
            value = parse_date(value)
        ws[coord] = value
        if coord[0] in WRAP_PREFIXES:
            wrap(ws, coord)

    wb.save(path)
    print(f"Saved {path} — tab {args.tab!r} from {src_name!r}")
    for coord in sorted(
        answers, key=lambda c: (c[0], int("".join(filter(str.isdigit, c)) or 0))
    ):
        print(f"  {coord.upper()}: {ws[str(coord).upper()].value!r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
