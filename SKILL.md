---
name: mytime-calender-booker
description: Full end-to-end MyTime time-booking automation for Outlook users. Fetches calendar events via COM automation (no browser, no login, no MFA), intelligently maps each meeting to a MyTime project and task using AI reasoning, then generates a filled Excel timecard automatically. Use this skill whenever the user mentions calendar, meetings, MyTime, time booking, or timecard — even if they don't use exact phrases. Triggers include "get my calendar", "fetch my meetings", "show my schedule", "sync my calendar", "start mytime booking", "book my meetings to mytime", "map my calendar to mytime", "confirm mytime booking", "review mytime mappings", "create timecard", "generate timecard", "book time", or any similar phrasing.
---

# mytime-calender-booker

End-to-end automation for booking Outlook calendar meetings into MyTime:
1. Fetch and review Outlook calendar events (PowerShell COM — no browser, no MFA)
2. Load MyTime projects and map each event to a project/task pair using AI reasoning
3. Generate a filled Excel timecard automatically as soon as mappings are confirmed

## Overview

Three phases, all implemented:

- **Phase 1:** Export calendar from Outlook → filter → present for review
- **Phase 2:** Load MyTime projects → agent maps events to project/task pairs → user confirms
- **Phase 3:** Mappings piped directly to Excel generator → file ready immediately

---

## Phase 1: Calendar Data Gathering

### Step 1 — Ask the user for filter preferences

Ask both questions at once, do not ask separately:

**Question 1 — Time range:**
> "Which time range would you like to fetch?
> (1) This week  (2) Today  (3) Custom — I'll specify start and end dates"

**Question 2 — Private meetings:**
> "Should I skip private meetings? (yes / no)"

Wait for both answers before proceeding.

If the user chose custom range, ask:
> "Please provide the start date and end date (format: YYYY-MM-DD)."

Once you have the answers, calculate concrete `StartDate` and `EndDate` strings (YYYY-MM-DD):
- **This week:** Monday of current week → Sunday of current week
- **Today:** today's date → today's date
- **Custom:** use the dates provided by the user

---

### Step 2 — Check Outlook is accessible

Check whether Outlook is running:

```bash
powershell -Command "if ((Get-Process OUTLOOK -ErrorAction SilentlyContinue) -or (Get-Process olk -ErrorAction SilentlyContinue)) { 'running' } else { 'not running' }"
```

- If `running`: proceed directly to Step 3.
- If `not running`: tell the user "Outlook isn't running — starting it now..." then proceed to Step 3. The export script handles starting it automatically (supports both classic `OUTLOOK.EXE` and new Outlook `olk.exe`).

---

### Step 3 — Export the calendar from Outlook

Run the export script with the calculated date range. Build the command based on user's answers:

```bash
# This week, include private:
powershell -ExecutionPolicy Bypass -File "D:\ai\custom-skills\mytime-calender-booker\scripts\export-calendar.ps1" -StartDate "2026-03-17" -EndDate "2026-03-23"

# This week, skip private:
powershell -ExecutionPolicy Bypass -File "D:\ai\custom-skills\mytime-calender-booker\scripts\export-calendar.ps1" -StartDate "2026-03-17" -EndDate "2026-03-23" -SkipPrivate

# Today, skip private:
powershell -ExecutionPolicy Bypass -File "D:\ai\custom-skills\mytime-calender-booker\scripts\export-calendar.ps1" -StartDate "2026-03-20" -EndDate "2026-03-20" -SkipPrivate

# Custom range:
powershell -ExecutionPolicy Bypass -File "D:\ai\custom-skills\mytime-calender-booker\scripts\export-calendar.ps1" -StartDate "2026-03-20" -EndDate "2026-03-27"
```

The script:
- Attaches to the running Outlook instance via COM (or starts Outlook if not running, then waits up to 60s)
- Exports the calendar to `%USERPROFILE%\.mytime-booker\calendar.ics`
- Outputs progress to stdout so you can see what's happening

If the script fails:
- `Could not find Outlook` → Outlook is not installed at the expected path. Ask the user to open Outlook manually and try again.
- `Export failed` → Show the error and ask the user to try again.
- `did not become ready within 60 seconds` → Outlook is taking too long to start. Ask the user to open Outlook manually, wait for it to fully load, then retry from Step 3.

---

### Step 4 — Parse and filter the exported ICS

Run the parser against the locally exported file:

```bash
# This week, include private:
python "D:\ai\custom-skills\mytime-calender-booker\scripts\parse-ics.py" --range this-week

# This week, skip private:
python "D:\ai\custom-skills\mytime-calender-booker\scripts\parse-ics.py" --range this-week --skip-private

# Today, skip private:
python "D:\ai\custom-skills\mytime-calender-booker\scripts\parse-ics.py" --range today --skip-private

# Custom range, skip private:
python "D:\ai\custom-skills\mytime-calender-booker\scripts\parse-ics.py" --range custom --start 2026-03-20 --end 2026-03-27 --skip-private
```

The `--range`, `--start`, `--end`, and `--skip-private` flags must match what was passed to the export script.

The parser output is a JSON array. Each event contains:
- `title`, `date`, `start`, `end`, `duration_hours` — core scheduling fields
- `description` — the actual meeting body text, cleaned of Teams join URLs and dial-in noise. This is the most valuable field for Phase 2 task mapping.
- `is_private`, `status`, `location`, `organizer`, `attendees`, `recurring`

**Note on descriptions:** Outlook writes two `DESCRIPTION` fields per event (real body first, then `"Reminder"`). The parser automatically keeps the longer/real one. If an event has no meaningful body, `description` will be `"Reminder"` or `null`.

If the script exits with an error:
- `ICS file not found` → the export script did not run or failed silently. Go back to Step 3.
- Any other error → show the error message and ask the user how to proceed.

---

### Step 5 — Present events for review

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

Then say:
> "Please review the list above. You can:
> - Say **ok** to confirm and hold these events for the next step
> - Say **remove #N** to exclude a specific event
> - Say **change range** to re-fetch with a different date range
> - Say **stop** to end here"

---

### Step 6 — Handle user review actions

**If user says "ok" or "confirm":**
Store the final filtered event list in context as `calendar_events`. The full JSON including `description` fields is available and will be used in Phase 2 to intelligently map events to MyTime projects and tasks. Tell the user:
> "Got it. I have [N] event(s) confirmed and ready. When you're ready to proceed to MyTime booking, just say the word — we'll handle project and task mapping in the next phase."

Do not proceed to Phase 2. Stop here and wait.

**If user says "remove #N":**
Remove the specified event(s) from the list, show the updated table, and ask for confirmation again.

**If user says "change range":**
Go back to Step 1.

**If user says "stop":**
> "Understood. The calendar data has been discarded. Let me know when you'd like to try again."

---

## Phase 2: MyTime Project/Task Mapping

Prerequisite: `calendar_events` must be in context from Phase 1. Do not start Phase 2 unless the user has confirmed their calendar events.

---

### Step 7 — Check if project data is fresh

Check whether `projects.json` exists:

```bash
powershell -Command "Test-Path \"$env:USERPROFILE\.mytime-booker\projects.json\""
```

Ask the user:
> "Have your MyTime projects changed since the last run? (yes / no)"

- **If `projects.json` does not exist:** treat as "yes" (first run / no cached data)
- **If user says "no":** load `projects.json` directly and skip to Step 9
- **If user says "yes":** proceed to Step 8

---

### Step 8 — Collect fresh MyTime projects

The MyTime `/my_projects` page requires JavaScript to load all tasks. The most reliable approach is a one-time HTML export.

Tell the user:
> "Please do the following in Edge (takes about 2 minutes):
> 1. Open your MyTime **my_projects** page in your browser
> 2. Click the **arrow icon** on every project to expand it and reveal all tasks
> 3. Save the page: press **Ctrl+S** → save it somewhere easy to find (e.g. Desktop)
> 4. Give me the file path"

Wait for the file path. Then parse it:

```bash
python "D:\ai\custom-skills\mytime-calender-booker\scripts\parse-projects.py" --file "C:\Users\MatthiasBigl\Desktop\My Time.html"
```

The script:
- Reads the saved HTML file
- Extracts all projects and tasks (name, ID, nickname, active dates)
- Saves to `%USERPROFILE%\.mytime-booker\projects.json`
- Outputs JSON to stdout

**If the script shows a warning about projects with 0 tasks:** those projects were collapsed in the browser when the page was saved. Tell the user which projects are affected and ask them to re-save the page after expanding those specific projects.

---

### Step 9 — Map calendar events to MyTime projects and tasks

Prerequisite: `calendar_events` must be in context from Phase 1.

Read `projects.json` to get the available projects and tasks.

For each event in `calendar_events`, find the best-matching (project, task) pair using these signals:

#### Signal 1 — Email domain match (strongest)

Extract all email addresses from the event's `organizer` and `attendees` fields. Strip the domain from each address.

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
- A project named `BAW-something` or containing "BAW" likely maps to events that mention "BAWAG" or "bawaggroup"
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

### Step 10 — Present mapping for review

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

### Step 11 — Handle mapping review actions

**If user says "ok" or "confirm":**

1. **Generate the Excel immediately** by piping the confirmed mappings as JSON directly to the script — no intermediate file needed. Build a JSON array from `booking_mappings` with these fields per event: `title`, `date`, `start`, `end`, `duration_hours`, `project_id`, `project_name`, `task_id`, `task_name`. **Do not include events that were skipped.**

```powershell
$json = @'
[
  {
    "title": "Team Sync",
    "date": "2026-03-24",
    "start": "09:00",
    "end": "10:00",
    "duration_hours": 1.0,
    "project_id": "12345",
    "project_name": "Customer Project A",
    "task_id": "01",
    "task_name": "Development"
  }
]
'@
$json | python "D:\ai\custom-skills\mytime-calender-booker\scripts\book-timecard.py" --output "$env:USERPROFILE\Downloads\timecard_output.xlsx"
```

Replace the example with the actual confirmed events.

2. **Present the output** (OK/UNMAPPED rows) as reported by the script and tell the user:
> "Your timecard has been saved to `Downloads\timecard_output.xlsx` — ready to upload to MyTime. If anything needs adjusting, say **pick #N** to remap an event or **skip #N** to remove one."

**If user says "pick #N":**
Present a searchable project+task menu for that event. The user picks or types a search term. Update the mapping and show the updated table. Ask for confirmation again.

**If user says "skip #N":**
Remove the event from the mapping list. Show the updated table. Ask for confirmation again.

**If user says "stop":**
> "Understood. The mappings have been discarded. Let me know when you'd like to try again."

---

## Phase 3: Final Timecard Review

The Excel is generated automatically after Step 11 confirmation. If the user requests changes via **pick #N** or **skip #N**:

### Step 12 — Handle post-generation adjustments

**"pick #N":**
1. Re-open `booking_mappings` in context.
2. For row N, present the list of available projects and tasks from `projects.json`.
3. Let the user choose.
4. Update the event in `booking_mappings`.
5. Re-pipe the updated JSON and re-run `book-timecard.py` (same stdin-pipe command as Step 11).
6. Show the updated output.

**"skip #N":** Remove the row from `booking_mappings`, re-pipe JSON, re-run `book-timecard.py`, show updated output.

---

## Requirements

- PowerShell (Windows) — for Outlook COM automation
- Node.js — no longer required (ICS parser rewritten in Python)
- Python 3 — for Phase 3 (`book-timecard.py`)
- `openpyxl` — install with: `pip install openpyxl`

## File structure reference

```
mytime-calender-booker/
  SKILL.md                            ← this file
  scripts/
    export-calendar.ps1               ← Outlook COM export (PowerShell)
    parse-ics.py                       ← ICS parser and filter (Python, stdlib)
    parse-projects.py                  ← MyTime HTML → projects.json (Python, stdlib)
    book-timecard.py                   ← Writes pre-mapped events into xlsx template (Python, openpyxl)
```

## Output files

| File | When written | Overwritten on next run |
|---|---|---|
| `%USERPROFILE%\.mytime-booker\calendar.ics` | Every Phase 1 export | Yes — always fresh |
| `%USERPROFILE%\.mytime-booker\projects.json` | When user confirms projects changed | Only when user says "yes" |
| `Downloads\timecard_output.xlsx` | After Step 11 confirmation | Yes — overwritten each confirmation |
