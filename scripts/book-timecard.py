import argparse
import csv
import os
import sys

import openpyxl

TEMPLATE_PATH = os.path.join(os.path.expanduser("~"), ".mytime-booker", "timecard_template.xlsx")
DEFAULT_OUTPUT = os.path.join(os.path.expanduser("~"), "Downloads", "timecard_output.xlsx")
DEFAULT_TYPE = "Normal -AT"

# CSV columns (must match what the agent writes)
COLUMNS = [
    "project_number",
    "project_name",
    "task_number",
    "task_name",
    "type",
    "date",
    "hours",
    "comment",
    "time_from",
    "time_to",
]


def load_csv(path):
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return list(reader)


def save_timecard(rows, output_path, template_path=None):
    tp = template_path or TEMPLATE_PATH
    if not os.path.exists(tp):
        print(f"[book-timecard] ERROR: template not found at: {tp}", file=sys.stderr)
        sys.exit(1)

    wb = openpyxl.load_workbook(tp)
    ws = wb["Timecard"]

    for row in rows:
        ws.append([
            row.get("project_number", ""),
            row.get("project_name", ""),
            row.get("task_number", ""),
            row.get("task_name", ""),
            row.get("type") or DEFAULT_TYPE,
            row.get("date", ""),
            row.get("hours", ""),
            row.get("comment", ""),
            row.get("time_from", ""),
            row.get("time_to", ""),
        ])

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    print(f"[book-timecard] Saved {len(rows)} row(s) to: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Append pre-mapped calendar events into a MyTime timecard xlsx"
    )
    parser.add_argument(
        "--events", required=True,
        help="Path to CSV file with confirmed, pre-mapped events"
    )
    parser.add_argument(
        "--template", default=TEMPLATE_PATH,
        help="Path to the blank template xlsx"
    )
    parser.add_argument(
        "--output", default=DEFAULT_OUTPUT,
        help="Output path for the filled xlsx"
    )
    args = parser.parse_args()

    if not os.path.exists(args.events):
        print(f"[book-timecard] ERROR: events file not found: {args.events}", file=sys.stderr)
        sys.exit(1)

    rows = load_csv(args.events)
    print(f"[book-timecard] Loaded {len(rows)} event(s)")

    if not rows:
        print("[book-timecard] No rows to book.", file=sys.stderr)
        sys.exit(1)

    for i, r in enumerate(rows, 1):
        status = "OK" if r.get("project_number") else "UNMAPPED"
        print(f"  {i}. [{status}] {r.get('date')} {r.get('time_from')}-{r.get('time_to')} | "
              f"{r.get('project_number')}/{r.get('task_number')} | {r.get('comment', '')[:50]}")

    save_timecard(rows, args.output, args.template)
    print(f"[book-timecard] Done: {args.output}")


if __name__ == "__main__":
    main()
