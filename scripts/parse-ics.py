#!/usr/bin/env python3
"""
parse-ics.py
Reads and parses a locally exported Outlook ICS calendar file.
Filters by date range and optionally skips private events.

Usage:
  python parse-ics.py --range this-week [--skip-private] [--file path/to/calendar.ics]
  python parse-ics.py --range today [--skip-private]
  python parse-ics.py --range custom --start 2026-03-20 --end 2026-03-27 [--skip-private]

The ICS file is produced by export-calendar.ps1 before calling this script.
Default file location: %USERPROFILE%\\.mytime-booker\\calendar.ics

Output: JSON array of events to stdout
"""

import argparse
import json
import os
import re
import sys
from datetime import date, datetime, timedelta


# ---------------------------------------------------------------------------
# CLI argument parsing
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(description="Parse Outlook ICS calendar export")
    parser.add_argument("--range", default="this-week", choices=["today", "this-week", "custom"])
    parser.add_argument("--skip-private", action="store_true")
    parser.add_argument("--start", help="Start date YYYY-MM-DD (required for --range custom)")
    parser.add_argument("--end", help="End date YYYY-MM-DD (required for --range custom)")
    parser.add_argument("--file", help="Path to .ics file (default: ~/.mytime-booker/calendar.ics)")
    return parser.parse_args()


# ---------------------------------------------------------------------------
# Date range calculation
# ---------------------------------------------------------------------------

def get_date_range(range_name, custom_start=None, custom_end=None):
    today = date.today()

    if range_name == "today":
        return today, today

    if range_name == "this-week":
        # Week starts Monday (weekday 0)
        monday = today - timedelta(days=today.weekday())
        sunday = monday + timedelta(days=6)
        return monday, sunday

    if range_name == "custom":
        if not custom_start or not custom_end:
            print("--range custom requires --start YYYY-MM-DD and --end YYYY-MM-DD", file=sys.stderr)
            sys.exit(1)
        try:
            return (
                datetime.strptime(custom_start, "%Y-%m-%d").date(),
                datetime.strptime(custom_end, "%Y-%m-%d").date(),
            )
        except ValueError as e:
            print(f"Invalid date format: {e}", file=sys.stderr)
            sys.exit(1)


# ---------------------------------------------------------------------------
# ICS line unfolding
# ---------------------------------------------------------------------------

def unfold_lines(raw):
    # ICS lines fold at 75 chars with CRLF + whitespace continuation
    return re.sub(r"\r?\n[ \t]", "", raw)


# ---------------------------------------------------------------------------
# ICS date/datetime parsing
# ---------------------------------------------------------------------------

def parse_ics_date(value):
    """Parse an ICS date or datetime value. Returns a date or datetime object."""
    if not value:
        return None

    # Strip TZID=...: prefix if present
    clean = re.sub(r"^TZID=[^:]+:", "", value)
    is_utc = clean.endswith("Z")
    d = clean.rstrip("Z")

    if len(d) == 8:
        # All-day: YYYYMMDD
        return datetime(int(d[0:4]), int(d[4:6]), int(d[6:8])).date()

    # Datetime: YYYYMMDDTHHmmss
    try:
        dt = datetime(
            int(d[0:4]), int(d[4:6]), int(d[6:8]),
            int(d[9:11]), int(d[11:13]), int(d[13:15])
        )
    except (ValueError, IndexError):
        return None

    return dt


# ---------------------------------------------------------------------------
# ICS property line parsing
# ---------------------------------------------------------------------------

def parse_property_line(line):
    """Returns (key, params_dict, value)."""
    colon_idx = line.find(":")
    if colon_idx == -1:
        return line, {}, ""

    key_part = line[:colon_idx]
    value = line[colon_idx + 1:]

    semi_idx = key_part.find(";")
    if semi_idx == -1:
        key = key_part
        param_str = ""
    else:
        key = key_part[:semi_idx]
        param_str = key_part[semi_idx + 1:]

    params = {}
    if param_str:
        for p in param_str.split(";"):
            if "=" in p:
                pk, pv = p.split("=", 1)
                params[pk] = pv
            else:
                params[p] = ""

    return key, params, value


# ---------------------------------------------------------------------------
# ICS parser
# ---------------------------------------------------------------------------

def parse_ics(raw):
    lines = unfold_lines(raw).splitlines()
    events = []
    current = None

    for line in lines:
        if line == "BEGIN:VEVENT":
            current = {}
            continue
        if line == "END:VEVENT":
            if current is not None:
                events.append(current)
            current = None
            continue
        if current is None:
            continue

        key, params, value = parse_property_line(line)

        if key == "SUMMARY":
            current["title"] = value.replace("\\,", ",").replace("\\n", "\n").strip()

        elif key == "DTSTART":
            current["dtstart"] = value
            tzid = params.get("TZID")
            full_val = f"TZID={tzid}:{value}" if tzid else value
            parsed = parse_ics_date(full_val)
            current["start"] = parsed
            current["all_day"] = isinstance(parsed, date) and not isinstance(parsed, datetime)

        elif key == "DTEND":
            current["dtend"] = value
            tzid = params.get("TZID")
            full_val = f"TZID={tzid}:{value}" if tzid else value
            current["end"] = parse_ics_date(full_val)

        elif key == "DURATION":
            current["duration_raw"] = value

        elif key == "CLASS":
            current["classification"] = value.upper()

        elif key == "STATUS":
            current["status"] = value.upper()

        elif key == "TRANSP":
            current["transp"] = value.upper()

        elif key == "LOCATION":
            current["location"] = value.replace("\\,", ",").strip()

        elif key == "DESCRIPTION":
            parsed = value.replace("\\n", "\n").replace("\\,", ",").strip()
            # Outlook writes two DESCRIPTION fields: real content first, then "Reminder"
            # Keep whichever is longer (the real one)
            existing = current.get("description", "")
            if len(parsed) > len(existing):
                current["description"] = parsed

        elif key == "ORGANIZER":
            current["organizer"] = re.sub(r"^mailto:", "", value, flags=re.IGNORECASE)

        elif key == "ATTENDEE":
            current.setdefault("attendees", []).append(
                re.sub(r"^mailto:", "", value, flags=re.IGNORECASE)
            )

        elif key == "UID":
            current["uid"] = value

        elif key == "RRULE":
            current["recurring"] = True
            current["rrule"] = value

    return events


# ---------------------------------------------------------------------------
# Duration calculation
# ---------------------------------------------------------------------------

def duration_hours(start, end):
    if not start or not end:
        return None
    if isinstance(start, date) and not isinstance(start, datetime):
        return None
    delta = end - start
    h = delta.total_seconds() / 3600
    return round(h, 2)


# ---------------------------------------------------------------------------
# Output formatting
# ---------------------------------------------------------------------------

def format_date(d):
    if d is None:
        return None
    if isinstance(d, datetime):
        return d.date().isoformat()
    return d.isoformat()


def format_time(d):
    if d is None or not isinstance(d, datetime):
        return None
    return d.strftime("%H:%M")


def clean_description(raw):
    if not raw:
        return None
    text = raw
    # Strip Microsoft Teams join URLs and safelinks
    text = re.sub(r"https?://eur\d+\.safelinks[^\s>]*", "", text)
    text = re.sub(r"https?://teams\.microsoft[^\s>]*", "", text)
    text = re.sub(r"<https?://[^>]*>", "", text)
    text = re.sub(r"Jetzt an der Besprechung teilnehmen\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"Microsoft Teams.*?Hilfe\?", "", text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r"Benötigen Sie Hilfe\?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"_{5,}", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = text.strip()
    return text if text else None


def format_event(ev):
    start = ev.get("start")
    end = ev.get("end")
    all_day = ev.get("all_day", False)
    is_private = ev.get("classification") in ("PRIVATE", "CONFIDENTIAL")

    return {
        "uid": ev.get("uid"),
        "title": ev.get("title", "(no title)"),
        "date": format_date(start),
        "start": None if all_day else format_time(start),
        "end": None if all_day else format_time(end),
        "duration_hours": None if all_day else duration_hours(start, end),
        "all_day": all_day,
        "is_private": is_private,
        "status": ev.get("status", "CONFIRMED"),
        "location": ev.get("location") or None,
        "description": clean_description(ev.get("description")),
        "organizer": ev.get("organizer") or None,
        "attendees": ev.get("attendees", []),
        "recurring": ev.get("recurring", False),
    }


# ---------------------------------------------------------------------------
# Deduplication
# ---------------------------------------------------------------------------

def data_score(ev):
    desc = ev.get("description", "")
    desc_score = len(desc) if desc and desc != "Reminder" else 0
    return (
        (2 if ev.get("title") and ev.get("title") != "(no title)" else 0)
        + (1 if ev.get("organizer") else 0)
        + len(ev.get("attendees") or [])
        + (10 if desc_score > 0 else 0)
    )


def deduplicate(events):
    seen = {}
    for ev in events:
        key = ev.get("uid") or f"{ev.get('dtstart')}-{ev.get('title')}"
        if key not in seen or data_score(ev) > data_score(seen[key]):
            seen[key] = ev
    return list(seen.values())


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def event_date(ev):
    s = ev.get("start")
    if isinstance(s, datetime):
        return s.date()
    return s


def main():
    args = parse_args()

    default_ics = os.path.join(
        os.environ.get("USERPROFILE") or os.environ.get("HOME") or ".",
        ".mytime-booker", "calendar.ics"
    )
    ics_path = args.file or default_ics

    start_date, end_date = get_date_range(args.range, args.start, args.end)

    if not os.path.exists(ics_path):
        print(f"ICS file not found at: {ics_path}", file=sys.stderr)
        print("Run export-calendar.ps1 first to generate it.", file=sys.stderr)
        sys.exit(1)

    with open(ics_path, encoding="utf-8", errors="replace") as f:
        raw = f.read()

    print(f"[parse-ics] Reading from: {ics_path}", file=sys.stderr)

    all_events = parse_ics(raw)
    deduped = deduplicate(all_events)

    filtered = []
    for ev in deduped:
        start = ev.get("start")
        if not start:
            continue
        if ev.get("status") == "CANCELLED":
            continue
        title = ev.get("title", "")
        if title.startswith("Canceled:") or not title or title == "(no title)":
            continue
        if args.skip_private and ev.get("classification") in ("PRIVATE", "CONFIDENTIAL"):
            continue
        ev_date = event_date(ev)
        if ev_date is None or not (start_date <= ev_date <= end_date):
            continue
        filtered.append(ev)

    # Sort by start time
    filtered.sort(key=lambda e: e.get("start") or date.min)

    output = [format_event(ev) for ev in filtered]
    print(json.dumps(output, indent=2, ensure_ascii=False))

    print(
        f"\n[parse-ics] {len(output)} event(s) found between {start_date} and {end_date}",
        file=sys.stderr
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Unexpected error: {e}", file=sys.stderr)
        sys.exit(1)
