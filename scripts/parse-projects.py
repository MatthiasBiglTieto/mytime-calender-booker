#!/usr/bin/env python3
"""
parse-projects.py
Parses a saved MyTime /my_projects HTML page into a structured TOON file.

Usage:
    python parse-projects.py --file "C:/path/to/My Time.html"
    python parse-projects.py --file "C:/path/to/My Time.html" --out "C:/path/to/projects.toon"

The output defaults to %USERPROFILE%/.mytime-booker/projects.toon

Dependencies: Python 3.x + toon_format
    pip install git+https://github.com/toon-format/toon-python.git
"""

import argparse
import os
import re
import sys
from datetime import datetime, timezone

from toon_format import encode as toon_encode
from html.parser import HTMLParser


class MyTimeProjectParser(HTMLParser):
    """
    Parses the MyTime /my_projects HTML.

    The page has a flat <table> where each project has:
      <tr class="project_row" id="project_{PID}">          ← project row
      <tr id="my_project_{PID}_my_tasks">                ← task section (wraps a nested <table>)
        <td><table><tbody>
          <tr class="task_row">...                        ← task rows inside nested table
          <tr class="task_row">...
        </tbody></table></td>
      </tr>

    We use a nesting counter to detect when we're inside the nested task table.
    """

    def __init__(self):
        super().__init__()
        self.projects = []
        self._project = None      # currently building project
        self._task = None          # currently building task
        self._tasks_pid = None     # project_id whose task section we're in

        # Nesting inside task section's nested <table>
        self._task_depth = 0

        # Project-row column counter
        self._proj_col = 0

        # Inline text capture (per-element, reset after use)
        self._buf = ""

    # ── Helpers ────────────────────────────────────────────────────────────────

    def _flush_buf(self):
        text = self._buf.strip()
        self._buf = ""
        return text

    def _find_proj(self, pid):
        for p in self.projects:
            if p["id"] == pid:
                return p
        return None

    # ── Tag handlers ─────────────────────────────────────────────────────────

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        classes = attrs.get("class", "")
        eid = attrs.get("id", "")
        name = attrs.get("name", "")
        val = attrs.get("value", "")

        # ── Project row ────────────────────────────────────────────────────────
        if tag == "tr" and "project_row" in classes:
            m = re.search(r"project_(\d+)", eid)
            if m:
                self._project = {
                    "id": m.group(1), "name": "", "active_dates": "",
                    "project_manager": None, "nickname": "",
                    "comment_required": False, "tasks": [],
                }
                self._proj_col = 0

        # Column counter inside project row
        if tag == "td" and self._project is not None and self._task is None:
            self._proj_col += 1

        # Project name — col 4 label: <label for="project_{PID}_selected_flag">
        if (tag == "label" and self._project and self._task is None
                and attrs.get("for", "").startswith("project_")
                and attrs.get("for", "").endswith("_selected_flag")):
            self._buf = ""

        # Project nickname: <input class="description" value="...">
        if (tag == "input" and self._project and self._task is None
                and "description" in classes and val):
            self._project["nickname"] = val

        # Project manager — col 6: <a href="mailto:...">
        if tag == "a" and self._project and self._task is None and self._proj_col == 6:
            self._buf = ""

        # Comment required — col 8: has checkmark (captured in handle_data)

        # ── Task section wrapper row ──────────────────────────────────────────
        if (tag == "tr" and eid.startswith("my_project_")
                and eid.endswith("_my_tasks")):
            m = re.search(r"my_project_(\d+)_my_tasks", eid)
            if m:
                self._tasks_pid = m.group(1)

        # Entering the nested task <table> inside the task section
        if tag == "table" and self._tasks_pid is not None:
            self._task_depth = 1

        # Inside task nested table
        if self._task_depth > 0:
            # Task row
            if tag == "tr" and "task_row" in classes:
                self._task = {"id": "", "name": "", "active_dates": ""}

            # Task ID hidden input: <input name="...[task_id]" value="{TID}">
            if tag == "input" and self._task and name.endswith("[task_id]"):
                self._task["id"] = val

            # Task name: <label for="task_{TID}_selected_flag">
            if (tag == "label" and self._task
                    and attrs.get("for", "").startswith("task_")):
                self._buf = ""

            # Task active dates cell
            if tag == "td" and self._task and "task_active" in classes:
                self._buf = ""

    def handle_endtag(self, tag):
        # ── Task table depth ──────────────────────────────────────────────────
        if tag == "table" and self._task_depth > 0:
            self._task_depth -= 1
            if self._task_depth == 0:
                self._tasks_pid = None

        # ── Project row ───────────────────────────────────────────────────────
        if tag == "tr" and self._project and self._task is None:
            text = self._flush_buf()
            if text and not self._project["name"]:
                self._project["name"] = text
            self.projects.append(self._project)
            self._project = None

        # Project name label
        if tag == "label" and self._project and self._task is None:
            self._project["name"] = self._flush_buf()

        # Project active dates (col 5) or PM (col 6)
        if tag == "td" and self._project and self._task is None:
            text = self._flush_buf()
            if not text:
                return
            if self._proj_col == 5:
                self._project["active_dates"] = text
            elif self._proj_col == 6:
                self._project["project_manager"] = text

        # ── Task row ─────────────────────────────────────────────────────────
        if tag == "tr" and self._task:
            proj = self._find_proj(self._tasks_pid)
            if proj and self._task["id"]:
                proj["tasks"].append(self._task)
            self._task = None

        # Task name label
        if tag == "label" and self._task:
            self._task["name"] = self._flush_buf()

        # Task active dates cell
        if tag == "td" and self._task:
            self._task["active_dates"] = self._flush_buf()

    def handle_data(self, data):
        self._buf += data

        # Comment required checkmark (col 8)
        if self._project and not self._task and self._proj_col == 8 and "\u2714" in data:
            self._project["comment_required"] = True


def parse_html(html_content):
    """Parse the MyTime HTML and return structured project data."""
    parser = MyTimeProjectParser()
    parser.feed(html_content)
    return parser.projects


def main():
    arg_parser = argparse.ArgumentParser(
        description="Parse MyTime /my_projects HTML into structured TOON"
    )
    arg_parser.add_argument(
        "--file",
        required=True,
        help="Path to the saved MyTime HTML file",
    )
    arg_parser.add_argument(
        "--out",
        default=os.path.join(
            os.environ.get("USERPROFILE", os.environ.get("HOME", ".")),
            ".mytime-booker",
            "projects.toon",
        ),
        help="Output path for projects.toon (default: ~/.mytime-booker/projects.toon)",
    )
    args = arg_parser.parse_args()

    # Read HTML
    if not os.path.exists(args.file):
        print(f"Error: File not found: {args.file}", file=sys.stderr)
        sys.exit(1)

    with open(args.file, "r", encoding="utf-8") as f:
        html_content = f.read()

    # Parse
    projects = parse_html(html_content)

    if not projects:
        print("Error: No projects found in the HTML. Is this the right page?", file=sys.stderr)
        sys.exit(1)

    # Count stats
    total_tasks = sum(len(p["tasks"]) for p in projects)
    projects_without_tasks = [p for p in projects if not p["tasks"]]

    # Build output
    output = {
        "scraped_at": datetime.now(timezone.utc).isoformat(),
        "source_file": os.path.abspath(args.file),
        "project_count": len(projects),
        "task_count": total_tasks,
        "projects": projects,
    }

    # Encode to TOON
    toon_output = toon_encode(output)

    # Ensure output directory exists
    out_dir = os.path.dirname(args.out)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir)

    # Write TOON
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(toon_output)
        f.write("\n")

    # Summary to stderr
    print(f"[parse-projects] Parsed {len(projects)} projects with {total_tasks} tasks", file=sys.stderr)
    if projects_without_tasks:
        print(
            f"[parse-projects] WARNING: {len(projects_without_tasks)} project(s) have no tasks (were they expanded in the browser?):",
            file=sys.stderr,
        )
        for p in projects_without_tasks:
            print(f"  - {p['name']} (id: {p['id']})", file=sys.stderr)

    # Also print TOON to stdout
    sys.stdout.write(toon_output)
    sys.stdout.write("\n")


if __name__ == "__main__":
    main()
