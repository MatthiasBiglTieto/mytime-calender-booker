---
name: mytime-calender-booker
description: Full end-to-end MyTime time-booking automation for Outlook users. Fetches calendar events via COM automation (no browser, no login, no MFA), intelligently maps each meeting to a MyTime project and task using AI reasoning, then generates a filled Excel timecard automatically. Use this skill whenever the user mentions calendar, meetings, MyTime, time booking, or timecard — even if they don't use exact phrases. Triggers include "get my calendar", "fetch my meetings", "show my schedule", "sync my calendar", "start mytime booking", "book my meetings to mytime", "map my calendar to mytime", "confirm mytime booking", "review mytime mappings", "create timecard", "generate timecard", "book time", or any similar phrasing.
---

# mytime-calender-booker

Windows-only. Requires Outlook (prefers new Outlook / olk.exe). All scripts handle their own path resolution — do not hardcode user home paths.

**Path convention:** `~` means the user's home directory throughout this document. When you need the actual absolute path (e.g. for the Write tool), resolve it with `python -c "import os; print(os.path.expanduser('~'))"`. Skill script paths use `$HOME/.agents/skills/...` which expands correctly in the PowerShell terminal.

**Workflow:** Phase 1 (calendar export) → Phase 2 (project mapping) → Phase 3 (Excel generation). Follow steps in order, but if the user says they already have data from earlier phases, check for existing files and skip ahead.

## Step 0 — Check dependencies

Before starting any phase, verify that required Python packages are installed. Run this once per session:

```bash
pip install openpyxl 2>nul && python -c "import toon_format" 2>nul || pip install git+https://github.com/toon-format/toon-python.git
```

If either install fails, tell the user what went wrong and stop.

---

## Phase 1: Calendar Data Gathering

### Step 1 — Confirm filter preferences with a smart default

Lead with an assumed default to minimise round-trips:

> "I'll fetch **this week's** calendar and **skip private meetings** — does that work?
> Or say: **today** / **custom range** / **include private**"

Wait for the user's reply before proceeding.

- If the user says **ok / yes / sure** (or anything that confirms): use this week, skip private.
- If the user says **today**: use today's date, skip private (unless they add "include private").
- If the user says **custom range**: ask for start and end dates (format: YYYY-MM-DD), then proceed.
- If the user says **include private**: use this week but include all events.
- If the user says **today include private** or any combination: apply both.

Once you have the answers, calculate concrete `StartDate` and `EndDate` strings (YYYY-MM-DD):
- **This week:** Monday of current week → Sunday of current week
- **Today:** today's date → today's date
- **Custom:** use the dates provided by the user

---

### Step 2 — Export the calendar from Outlook

Run the export script with the calculated date range. Build the command based on user's answers:

```bash
# This week, include private:
powershell -ExecutionPolicy Bypass -File "$HOME/.agents/skills/mytime-calender-booker/scripts/export-calendar.ps1" -StartDate "2026-03-17" -EndDate "2026-03-23"

# This week, skip private:
powershell -ExecutionPolicy Bypass -File "$HOME/.agents/skills/mytime-calender-booker/scripts/export-calendar.ps1" -StartDate "2026-03-17" -EndDate "2026-03-23" -SkipPrivate

# Today, skip private:
powershell -ExecutionPolicy Bypass -File "$HOME/.agents/skills/mytime-calender-booker/scripts/export-calendar.ps1" -StartDate "2026-03-20" -EndDate "2026-03-20" -SkipPrivate

# Custom range:
powershell -ExecutionPolicy Bypass -File "$HOME/.agents/skills/mytime-calender-booker/scripts/export-calendar.ps1" -StartDate "2026-03-20" -EndDate "2026-03-27"
```

The script:
- Attaches to the running Outlook instance via COM, or starts it automatically if not running — no pre-check needed (handles both classic `OUTLOOK.EXE` and new `olk.exe`)
- Exports the calendar to `~\.mytime-booker\calendar.ics`
- Outputs progress to stdout so you can see what's happening

If the script fails:
- `Could not find Outlook` → Outlook is not installed at the expected path. Ask the user to open Outlook manually and try again.
- `Export failed` → Show the error and ask the user to try again.
- `did not become ready within 60 seconds` → Outlook is taking too long to start. Ask the user to open Outlook manually, wait for it to fully load, then retry from Step 2.

### Step 2b — Parse and filter the exported ICS

Run this command in the SAME turn as Step 2 — do not wait for user input between export and parse, because the user already confirmed their preferences and there is nothing new to ask.

```bash
# This week, include private:
python "$HOME/.agents/skills/mytime-calender-booker/scripts/parse-ics.py" --range this-week

# This week, skip private:
python "$HOME/.agents/skills/mytime-calender-booker/scripts/parse-ics.py" --range this-week --skip-private

# Today, skip private:
python "$HOME/.agents/skills/mytime-calender-booker/scripts/parse-ics.py" --range today --skip-private

# Custom range, skip private:
python "$HOME/.agents/skills/mytime-calender-booker/scripts/parse-ics.py" --range custom --start 2026-03-20 --end 2026-03-27 --skip-private
```

The `--range`, `--start`, `--end`, and `--skip-private` flags must match what was passed to the export script.

The parser output is a TOON array (a JSON-compatible text format — read it as JSON). Each event contains:
- `title`, `date`, `start`, `end`, `duration_hours` — core scheduling fields
- `description` — the actual meeting body text, truncated at the first Teams/phone boilerplate line (meeting IDs, dial-in numbers, etc.) so legal disclaimers are never included. If no meaningful content remains, `description` will be `null`.
- `is_private`, `location`, `organizer`, `attendee_domains` (unique email domains of attendees, e.g. `["acme-corp.com", "contoso.com"]`), `recurring`

If the parse script exits with an error:
- `ICS file not found` → the export script did not run or failed silently. Retry from Step 2.
- Any other error → show the error message and ask the user how to proceed.

---

### Step 3 — Present events for review

Parse the JSON output and present it as a clean, readable table. Include a short description preview (first 60 chars) where available:

```
Here are your meetings for this week (5 events):

#  Date        Day        Time           Duration  Title                              Description
──────────────────────────────────────────────────────────────────────────────────────────────────────────
1  2026-03-23  Monday     09:00–10:00    1.0h      Weekly Sync                        Agenda: sprint review...
2  2026-03-23  Monday     14:00–14:30    0.5h      1:1 with Sarah                     —
3  2026-03-24  Tuesday    (all day)      —         Company Holiday                    —
4  2026-03-25  Wednesday  10:00–11:00    1.0h      Sprint Planning                    Please prepare your PBI...
5  2026-03-26  Thursday   15:00–15:30    0.5h      Retrospective                      —
```

Rules for the description column:
- If `description` is null, `"Reminder"`, or empty → show `—`
- Otherwise show first 60 characters followed by `...` if truncated
- Strip any remaining `\n` in the preview (replace with space)

**Personal event detection:** After presenting the table, scan for events that look like personal/non-work appointments. Flag any event where ALL of the following are true:
- `organizer` is null
- `attendee_domains` is empty
- `location` is null or not a meeting URL (no "teams", "zoom", "meet", etc.)
- Title does NOT contain work-related terms like 'focus time', 'blocked', 'prep', 'no meeting', or 'lunch'

If any are found, proactively call them out and suggest removal:
> "Events #N, #M look like personal appointments (no organizer, no attendees, no meeting link) — want me to remove them?"

**All-day events:** Events like holidays, out-of-office, or all-day blocks typically should not be booked to projects. Proactively suggest skipping them:
> "Event #N is an all-day event ([title]) — want me to skip it for booking?"

Then say:
> "Please review the list above. You can:
> - Say **ok** to confirm and proceed to project mapping
> - Say **remove #N** to exclude a specific event
> - Say **change range** to re-fetch with a different date range
> - Say **stop** to end here"

---

### Step 4 — Handle user review actions

**If user says "ok" or "confirm":**
The confirmed event list (full JSON including descriptions) carries forward into Phase 2 for mapping. Immediately offer to continue:
> "Got it — [N] event(s) confirmed. **Ready to map them to MyTime projects now?** (yes / not yet)"

- If the user says **yes**: proceed directly to Phase 2 (Step 5).
- If the user says **not yet / later**: stop and wait.

**If user says "remove #N":**
Remove the specified event(s) from the list, show the updated table, and ask for confirmation again.

**If user says "change range":**
Go back to Step 1.

**If user says "stop":**
> "Understood. The calendar data has been discarded. Let me know when you'd like to try again."

---

## Phase 2: MyTime Project/Task Mapping

Prerequisite: the user must have confirmed their calendar events in Phase 1 before starting Phase 2.

---

### Step 5 — Check if project data is fresh

Check whether `projects.toon` exists and how old it is using Python (no PowerShell needed):

```bash
python -c "import os, pathlib, datetime; p = pathlib.Path(os.path.expanduser('~')) / '.mytime-booker' / 'projects.toon'; print('not found' if not p.exists() else str(round((datetime.datetime.now() - datetime.datetime.fromtimestamp(p.stat().st_mtime)).total_seconds() / 3600, 1)) + ' hour(s) ago')"
```

- **If `not found`:** treat as first run — proceed directly to Step 6.
- **If an age is returned:** Ask the user:
  > "Your MyTime projects were last refreshed **[age]**. Use cached data or refresh?"
  > (cached / refresh)
  - **cached:** read `projects.toon` with the Read tool at `~\.mytime-booker\projects.toon` (resolve the absolute home path with `python -c "import os; print(os.path.expanduser('~'))"` if needed) and proceed to Step 7
  - **refresh:** proceed to Step 6

---

### Step 6 — Collect fresh MyTime projects

The MyTime `/my_projects` page requires JavaScript to load all tasks. The most reliable approach is a one-time HTML export.

Tell the user:
> "Please do the following in Edge (takes about 2 minutes):
> 1. Open your MyTime **my_projects** page in your browser
> 2. Click the **arrow icon** on every project to expand it and reveal all tasks
> 3. Save the page: press **Ctrl+S** → save it somewhere easy to find (e.g. Desktop)
> 4. Give me the file path"

Wait for the file path. Then parse it:

```bash
python "$HOME/.agents/skills/mytime-calender-booker/scripts/parse-projects.py" --file "<path given by user>"
```

The script:
- Reads the saved HTML file
- Extracts all projects and tasks (name, ID, nickname, active dates)
- Saves to `~\.mytime-booker\projects.toon`
- Outputs JSON to stdout

**If the script shows a warning about projects with 0 tasks:** those projects were collapsed in the browser when the page was saved. Tell the user which projects are affected and ask them to re-save the page after expanding those specific projects.

---

### Step 7 — Map calendar events to MyTime projects and tasks

Read `projects.toon` to get the available projects and tasks.

For each confirmed event from Phase 1, find the best-matching (project, task) pair using these signals:

**Priority order: 1 > 2 > 3 > 4 > 5 > 6.** When signals conflict, the higher-numbered signal loses. Domain match is the strongest differentiator.

#### Signal 1 — Email domain match (strongest)

Extract all unique email domains from the event's `organizer` and `attendee_domains` fields.

- If all or most attendees share an external domain (e.g. `@customer.com`): prefer projects whose name or project manager suggests that customer. Do NOT match internal projects.
- If attendees are mostly internal (e.g. `@yourcompany.com`): match internal projects.
- The presence of an external domain is a strong filter — a meeting with `@external-client.com` attendees should almost never map to an internal project.

#### Signal 2 — Keyword match on project name

Compare the event `title` and `description` against project names. Look for shared words, acronyms, or project codes.

#### Signal 3 — Keyword match on task name

Compare the event `title` and `description` against task names. Look for task codes or descriptive phrases.

#### Signal 4 — Project number match

If the event `description` mentions a numeric project code (e.g. "project 12345" or "12345 -"), match it against project IDs/names.

#### Signal 5 — Naming conventions and abbreviations

Look at how project and task names are structured and use that to reason about matches:
- A project named `ACM-something` or containing "ACM" likely maps to events that mention "Acme" or "acme-corp"
- A task named "Cloud Support" likely matches stand-up, Wartung, or maintenance meetings
- A project/task containing "chapter" or "chapter work" likely matches chapter or webdev meetings
- A task named "Training" or "Competence" likely matches Copilot, GitHub Copilot, or AI-agent meetings

These are **soft hints to guide your reasoning**, not hard rules. Use your judgement — if other signals point elsewhere, prefer them.

#### Signal 6 — Default fallback

If no confident match is found, leave the event unmapped and let the user choose manually.

**Confidence levels:**
- `high`: email domain strongly narrows the project AND keyword/description confirms it
- `medium`: keyword or description matches but no strong domain signal
- `unmapped`: no confident match found

---

### Step 8 — Present mapping for review

Present all mappings as a table:

```
Here are the project/task mappings for your meetings:

#  Date        Time      Event                           Project                     Task               Conf
─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
1  2026-03-16  09:00    Team Sync                       12345 - Customer Project A  01 - Development     high
2  2026-03-16  13:00    Project Kickoff                  12345 - Customer Project A  02 - Planning        high
3  2026-03-18  10:00    Sprint Review                    67890 - Internal Project B  01 - Meetings        medium
4  2026-03-18  11:00    Client Call                      67890 - Internal Project B  ?                    unmapped
5  2026-03-19  14:00    Weekly Chapter Meeting            11111 - Chapter Work         03 - Chapter Work     high
```

Show the full project and task names — let the user see exactly what was matched. Truncate only the middle of long names if the table gets too wide. For `unmapped` events, show `?` for project and task.

Then say:
> "Please review the mappings above. You can:
> - Say **ok** to confirm all and proceed to booking
> - Say **pick #N** to choose a different project/task for a specific event
> - Say **skip #N** to exclude an event from booking
> - Say **stop** to end here"

---

### Step 9 — Handle mapping review actions

**If user says "ok" or "confirm":**

1. **Generate the Excel immediately** using the `Write` tool + the script defaults. Two steps:

**Step A — write `bookings.csv`:** Use the Write tool to create the file at `~\.mytime-booker\bookings.csv`. To get the absolute path, run `python -c "import os; print(os.path.expanduser('~'))"` first.

The CSV has exactly 10 columns matching the timecard template. Do not include skipped events. Pre-split project and task names before writing. Split on the FIRST ` - ` (space-dash-space) only, because names like `"ACM-CC-Cloud-Ops"` contain extra dashes that are part of the name. Everything before the first ` - ` is the number, everything after is the name:
- `"295189 - ACM-CC-Cloud-Ops"` → `project_number=295189`, `project_name=ACM-CC-Cloud-Ops`
- `"61.2 - Cloud support"` → `task_number=61.2`, `task_name=Cloud support`

`type` is always `Normal -AT` for regular meetings. Example:

```csv
project_number,project_name,task_number,task_name,type,date,hours,comment,time_from,time_to
295189,ACM-CC-Cloud-Ops,61.2,Cloud support,Normal -AT,2026-03-24,1.0,Team Sync,09:00,10:00
291648,CE DevActive Delivery,07,Chapter Work,Normal -AT,2026-03-26,1.0,Weekly WebDev 2026,14:00,15:00
```

**Step B — generate the Excel** (uses built-in defaults for input/output paths, no args needed):
```bash
python "$HOME/.agents/skills/mytime-calender-booker/scripts/book-timecard.py"
```

2. **Present the output** (OK/UNMAPPED rows) as reported by the script and tell the user:
> "Your timecard has been saved to `Downloads\timecard_output.xlsx` — ready to upload to MyTime. If anything needs adjusting, say **pick #N** to remap an event or **skip #N** to remove one."

**If user says "pick #N":**
Present a searchable project+task menu for that event. The user picks or types a search term. Update the mapping and show the updated table. Ask for confirmation again.

**If user says "skip #N":**
Remove the event from the mapping list. Show the updated table. Ask for confirmation again.

**If user says "stop":**
> "Understood. The mappings have been discarded. Let me know when you'd like to try again."

---

## File structure reference

```
mytime-calender-booker/
  SKILL.md                            ← this file
  assets/
    timecard_template.xlsx             ← blank MyTime timecard template (ships with skill)
  scripts/
    export-calendar.ps1               ← Outlook COM export (PowerShell, prefers new Outlook)
    parse-ics.py                       ← ICS parser and filter (Python, stdlib)
    parse-projects.py                  ← MyTime HTML → projects.toon (Python, toon_format)
    book-timecard.py                   ← Writes pre-mapped events into xlsx template (Python, openpyxl)
```

The Excel template (`assets/timecard_template.xlsx`) is bundled with the skill and found automatically by `book-timecard.py` — no manual setup needed.

## Output files

| File | When written | Overwritten on next run |
|---|---|---|
| `~\.mytime-booker\calendar.ics` | Every Phase 1 export | Yes — always fresh |
| `~\.mytime-booker\projects.toon` | When user confirms projects changed | Only when user says "yes" |
| `~\.mytime-booker\bookings.csv` | After Step 9 confirmation | Yes — overwritten each confirmation |
| `~\Downloads\timecard_output.xlsx` | After Step 9 confirmation | Yes — overwritten each confirmation |
