# mytime-calender-booker

> An AI agent skill that automates the full MyTime timecard booking workflow — from your Outlook calendar to a ready-to-import Excel file, with no browser, no login, and no manual data entry.

---

## What it does

1. **Fetches your Outlook calendar** via COM automation (no browser, no MFA, works offline)
2. **Maps each meeting to a MyTime project and task** using AI reasoning (email domain, keywords, attendees)
3. **Generates a filled Excel timecard** ready to upload to MyTime

Supports both **classic Outlook** (`OUTLOOK.EXE`) and **new modern Outlook** (`olk.exe`).

---

## Requirements

| Dependency | Purpose |
|---|---|
| Windows + PowerShell 5.1+ | Outlook COM automation |
| Microsoft Outlook (classic or new) | Calendar data source |
| Python 3 + `openpyxl` | ICS parsing, project mapping, Excel generation |

Install the Python dependency:
```bash
pip install openpyxl
```

---

## Installation

Install via the [skills CLI](https://skills.sh):

```bash
# Install globally for all agents
skills add MatthiasBiglTieto/mytime-calender-booker -g --all

# Or for a specific agent only
skills add MatthiasBiglTieto/mytime-calender-booker --agent copilot
```

---

## How to use

Just talk to your AI agent naturally:

> "Book my timecard"
> "Generate my timecard for this week"
> "Map my calendar meetings to MyTime"

The skill triggers automatically on any of these phrases and guides you through three phases:

---

## Workflow

### Phase 1 — Calendar export

The agent asks:
- **Time range**: this week / today / custom range
- **Skip private meetings**: yes / no

It then connects to Outlook via COM (starts it automatically if not running), exports the calendar to a local ICS file, and presents your meetings as a reviewed table:

```
#  Date        Day        Time           Dur    Title
1  2026-03-23  Monday     13:00–13:45    0.75h  Genesys Move 2 Cloud - Stand-up
2  2026-03-24  Tuesday    10:30–11:30    1.0h   Outbound Requirements Meeting
3  2026-03-26  Thursday   14:00–15:00    1.0h   Weekly WebDev Chapter
```

You can remove individual events before proceeding.

---

### Phase 2 — AI project/task mapping

The agent loads your MyTime projects from a cached `projects.json` (or re-exports from the MyTime HTML page if needed).

For each meeting it finds the best-matching **(project, task)** pair using:
- **Email domain** of organizer/attendees → identifies the customer project
- **Keyword matching** against project and task names
- **Naming conventions** (e.g. `BAW-*` = BAWAG, `chapter` = Chapter Work task)

Results are shown as a mapping table:

```
#  Date        Time      Event                        Project               Task                  Conf
1  2026-03-23  13:00     Genesys Stand-up + Wartung   295189 BAW-CC-Cloud   61.2 Cloud support    high
2  2026-03-24  10:30     Outbound requirements         295189 BAW-CC-Cloud   70.5 Outbound dev     medium
3  2026-03-26  14:00     Weekly WebDev 2026            291648 COMPDevActive  07 Chapter Work       high
```

You can pick a different project/task for any row before confirming.

---

### Phase 3 — Excel generation

Once confirmed, the agent immediately generates:

```
Downloads\timecard_output.xlsx
```

The Excel format exactly matches what MyTime expects for import:

| Column | Example |
|---|---|
| Project number | `295189` |
| Project name | `BAW-CC-Cloud-OU216` |
| Task number | `61.2` |
| Task name | `Cloud support` |
| Type | `Normal -AT` |
| Date | `2026-03-24` |
| Hours | `1.0` |
| Comment | meeting title |
| Time from | `10:30` |
| Time to | `11:30` |

---

## File structure

```
mytime-calender-booker/
  SKILL.md                   ← skill definition (read by the AI agent)
  README.md                  ← this file
  scripts/
    export-calendar.ps1      ← Outlook COM export (PowerShell, classic + new Outlook)
    parse-ics.py             ← ICS parser and date filter (Python, stdlib)
    parse-projects.py        ← MyTime HTML → projects.json (Python)
    book-timecard.py         ← confirmed mappings → timecard.xlsx (Python + openpyxl)
  config/
    config.example.json      ← optional config template
```

### Output files

| File | Location | Overwritten |
|---|---|---|
| `calendar.ics` | `%USERPROFILE%\.mytime-booker\` | Every run |
| `projects.json` | `%USERPROFILE%\.mytime-booker\` | Only when you say projects changed |
| `timecard_output.xlsx` | `%USERPROFILE%\Downloads\` | Every confirmed booking |

---

## MyTime project setup (first run only)

The first time you use the skill, it needs to learn your MyTime projects:

1. Open your MyTime **my_projects** page in Edge
2. Click the arrow icon on every project to expand all tasks
3. Press **Ctrl+S** to save the page (e.g. to Desktop)
4. Tell the agent the file path

The agent parses the HTML and caches the result to `projects.json` for all future runs.

---

## Supported Outlook versions

| Outlook | Process | COM method |
|---|---|---|
| Classic Outlook (Office 365 / 2019 / 2021) | `OUTLOOK.EXE` | `GetActiveObject` |
| New modern Outlook (Windows Store) | `olk.exe` | `New-Object -ComObject` |

Both are detected automatically. If neither is running, the skill starts whichever is installed.

---

## License

MIT
