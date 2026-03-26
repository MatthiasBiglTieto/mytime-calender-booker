---
name: mytime-calender-booker
description: Full end-to-end MyTime time-booking automation for Outlook users. Fetches calendar events via COM automation (no browser, no login, no MFA), intelligently maps each meeting to a MyTime project and task using AI reasoning, then generates a filled Excel timecard automatically. Use this skill whenever the user mentions calendar, meetings, MyTime, time booking, or timecard — even if they don't use exact phrases. Triggers include "get my calendar", "fetch my meetings", "show my schedule", "sync my calendar", "start mytime booking", "book my meetings to mytime", "map my calendar to mytime", "confirm mytime booking", "review mytime mappings", "create timecard", "generate timecard", "book time", or any similar phrasing.
---
# mytime-calender-booker

Windows-only. Requires Outlook.

**Path convention:** Use `$env:USERPROFILE` for absolute path resolution in Windows environments.

**Initialization:**
Execute directory change prior to script execution. Use relative paths for all subsequent operations.
```powershell
Set-Location -Path "$env:USERPROFILE\.agents\skills\mytime-calender-booker"
```

## Phase 1: Calendar Data Gathering

1. Execute dependency verification:
```powershell
pip install openpyxl 2>nul; pip install git+[https://github.com/toon-format/toon-python.git](https://github.com/toon-format/toon-python.git)
```
2. Ask the user for their date range and filter preferences explicitly. State: "Which date range should I use for your timecard? (e.g., this week, today). Would you like to skip private meetings?" Await response.
3. Calculate `StartDate` and `EndDate` strings based on user input. Execute calendar export:
```powershell
powershell -ExecutionPolicy Bypass -File ".\scripts\export-calendar.ps1" -StartDate "<YYYY-MM-DD>" -EndDate "<YYYY-MM-DD>" <optional:-SkipPrivate>
```
4. Execute ICS parsing:
```powershell
python ".\scripts\parse-ics.py" --range <type> --start <YYYY-MM-DD> --end <YYYY-MM-DD> <optional:--skip-private>
```
5. Render the output as a Markdown table. Proactively flag all-day or personal events (no organizer/attendees, no meeting links). Ask the user: "Please review the events. Reply 'ok' to proceed, or specify events to remove." Await response.
6. Apply any user-requested removals or modifications to the internal dataset.

## Phase 2: MyTime Project/Task Mapping

1. Verify project data freshness:
```powershell
python -c "import os, pathlib, datetime; p = pathlib.Path.home() / '.mytime-booker' / 'projects.toon'; print('missing' if not p.exists() else str(round((datetime.datetime.now() - datetime.datetime.fromtimestamp(p.stat().st_mtime)).total_seconds() / 3600, 1)) + ' hour(s) ago')"
```
2. If output is `missing` or the user requests a refresh, instruct the user to export the MyTime HTML page and provide the file path. Await response. Execute parsing with the provided path:
```powershell
python ".\scripts\parse-projects.py" --file "<user-provided-path>"
```
3. Read `projects.toon`. Map confirmed Phase 1 events to projects and tasks from this file using strict priority reasoning: Email domain match (strongest) > Project name keyword > Task name keyword > Project number match > Naming conventions. THIS MUST BE DONE BY YOU NO SCRIPTS. USE AI REASONING TO FIND BEST MATCH
4. Render the mapping as a Markdown table. Include full project/task names and confidence levels (`high`, `medium`, `unmapped`). Ask the user: "Review mapping. Reply 'ok' to confirm, 'pick [number]' to override, or 'skip [number]' to exclude." Await response.
5. If the user overrides, present a searchable project/task list, await their selection, update the dataset, and present the mapping table again.

## Phase 3: Timecard Generation

1. Format the final dataset. Extract project numbers and task numbers by splitting strings at the first ` - ` sequence.
2. Write `bookings.csv` to `$env:USERPROFILE\.mytime-booker\bookings.csv`. Enforce strict columns: `project_number,project_name,task_number,task_name,type,date,hours,comment,time_from,time_to`. Set type to `Normal -AT`.
3. Execute Excel generation script:
```powershell
python ".\scripts\book-timecard.py"
```
4. Output completion status. Inform the user the timecard is located at `$env:USERPROFILE\Downloads\timecard_output.xlsx`.
