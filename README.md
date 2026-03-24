# mytime-calender-booker

An AI agent skill that automates the full MyTime timecard booking workflow — from Outlook calendar to a ready-to-import Excel file.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Windows 10/11 | Required for Outlook COM automation |
| PowerShell 5.1+ | Included with Windows |
| Microsoft Outlook | Classic (`OUTLOOK.EXE`) or new (`olk.exe`) — must be installed |
| Python 3.9+ | [python.org](https://www.python.org/downloads/) |
| `openpyxl` | `pip install openpyxl` |
| `toon_format` | `pip install git+https://github.com/toon-format/toon-python.git` |

---

## Installation

### 1. Install the skill

```bash
skills add MatthiasBiglTieto/mytime-calender-booker
```

During installation, select **Global** scope and **Symlink** method so updates apply automatically.

### 2. Install Python dependencies

```bash
pip install openpyxl
pip install git+https://github.com/toon-format/toon-python.git
```

### 3. Set up your MyTime projects (first run only)

The skill needs a local copy of your MyTime project list to map meetings:

1. Open your MyTime **my_projects** page in a browser
2. Click the **expand arrow** on every project to reveal all tasks
3. Save the page: **Ctrl+S** → save anywhere (e.g. Desktop)
4. When the skill asks for your project file, provide the saved path

The skill parses the HTML and caches the result to `%USERPROFILE%\.mytime-booker\projects.toon`. You only need to repeat this step when your project assignments change.

---

## Usage

Trigger the skill by talking to your AI agent naturally:

```
"Book my timecard"
"Generate my timecard for this week"
"Map my calendar meetings to MyTime"
```

The skill walks you through three phases:

**Phase 1 — Calendar export**
Connects to Outlook via COM (no browser, no MFA), exports your calendar, and presents meetings for review. Defaults to **this week, skip private** — just confirm or say "today", "custom range", or "include private" to adjust.

**Phase 2 — Project/task mapping**
AI maps each meeting to a MyTime project and task using email domains, attendee lists, keywords, and naming conventions. You review the table and can override any mapping before confirming.

**Phase 3 — Excel generation**
Confirmed mappings are written directly to `Downloads\timecard_output.xlsx` in the format MyTime expects for import.

---

## Output

| File | Location | When written |
|---|---|---|
| `calendar.ics` | `%USERPROFILE%\.mytime-booker\` | Every export |
| `projects.toon` | `%USERPROFILE%\.mytime-booker\` | First run / when projects change |
| `timecard_output.xlsx` | `%USERPROFILE%\Downloads\` | After each confirmed booking |

---

## File structure

```
mytime-calender-booker/
  SKILL.md                   ← skill instructions (read by the AI agent)
  README.md                  ← this file
  scripts/
    export-calendar.ps1      ← Outlook COM export (PowerShell)
    parse-ics.py             ← ICS parser and date filter (Python, stdlib)
    parse-projects.py        ← MyTime HTML → projects.toon (Python, toon_format)
    book-timecard.py         ← mappings → timecard.xlsx (Python, openpyxl)
```

---

## Supported Outlook versions

| Version | Process name | Detection |
|---|---|---|
| Classic Outlook (Office 365 / 2019 / 2021) | `OUTLOOK.EXE` | `GetActiveObject` |
| New modern Outlook (Windows Store app) | `olk.exe` | `New-Object -ComObject` |

Both are detected automatically. If neither is running when the skill starts, it launches whichever is installed.

---

## Troubleshooting

**Outlook isn't found / export fails**
Make sure Outlook is installed and you can open it manually. If using the new Outlook app, ensure it has finished loading before retrying.

**Projects with 0 tasks after parsing**
Some projects were collapsed when you saved the page. Re-open MyTime, expand those specific projects, re-save, and provide the new file path.

**Excel not opening in MyTime**
Verify the column order matches exactly what MyTime expects: Project number, Project name, Task number, Task name, Type, Date, Hours, Comment, Time from, Time to.


