<#
.SYNOPSIS
    Exports Outlook calendar events to an ICS file using COM automation.

.DESCRIPTION
    Connects to the local Outlook client via COM (no browser, no auth needed).
    If Outlook is not running, starts it automatically and waits for it to be ready.
    Exports the calendar for the specified date range to a local ICS file.

.PARAMETER StartDate
    Start date of the export range (YYYY-MM-DD). Required.

.PARAMETER EndDate
    End date of the export range (YYYY-MM-DD). Required.

.PARAMETER SkipPrivate
    If set, private/confidential events are excluded from the export.

.PARAMETER OutFile
    Path to write the ICS file. Defaults to %USERPROFILE%\.mytime-booker\calendar.ics

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File export-calendar.ps1 -StartDate "2026-03-17" -EndDate "2026-03-23"
    powershell -ExecutionPolicy Bypass -File export-calendar.ps1 -StartDate "2026-03-17" -EndDate "2026-03-23" -SkipPrivate
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$StartDate,

    [Parameter(Mandatory=$true)]
    [string]$EndDate,

    [switch]$SkipPrivate,

    [string]$OutFile = "$env:USERPROFILE\.mytime-booker\calendar.ics"
)

$ErrorActionPreference = "Stop"

$classicExe = "C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE"
$modernExe  = "$env:LOCALAPPDATA\Microsoft\WindowsApps\olk.exe"

# ---------------------------------------------------------------------------
# Helper: detect which Outlook variant is running
# Returns "classic", "modern", or $null
# ---------------------------------------------------------------------------
function Get-OutlookVariant {
    if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { return "classic" }
    if (Get-Process olk     -ErrorAction SilentlyContinue) { return "modern"  }
    return $null
}

# ---------------------------------------------------------------------------
# Helper: get or start Outlook COM object
#
# Classic Outlook (OUTLOOK.EXE):
#   Uses GetActiveObject to attach to the running instance. This is the proper
#   way for classic Outlook and works in Windows PowerShell 5.1 (.NET Framework).
#
# New Outlook (olk.exe):
#   GetActiveObject is not supported; uses New-Object -ComObject instead, which
#   is how the new Outlook registers its COM interface.
# ---------------------------------------------------------------------------
function Get-OutlookCOM {
    $variant = Get-OutlookVariant

    if ($variant -eq "classic") {
        Write-Host "[export-calendar] Classic Outlook is running - attaching via GetActiveObject."
        try {
            return [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {
            Write-Warning "[export-calendar] GetActiveObject failed, falling back to New-Object COM."
            return New-Object -ComObject Outlook.Application
        }
    }

    if ($variant -eq "modern") {
        Write-Host "[export-calendar] New Outlook (olk) is running - attaching via New-Object COM."
        return New-Object -ComObject Outlook.Application
    }

    # Neither is running — start whichever is installed
    Write-Host "[export-calendar] Outlook is not running. Starting it (preferring new Outlook)..."

    $exeToStart = $null
    if (Test-Path $modernExe) {
        $exeToStart = $modernExe
    } elseif (Test-Path $classicExe) {
        $exeToStart = $classicExe
    } else {
        # Fallback: registry lookup for classic Outlook
        try {
            $regPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" -ErrorAction Stop).'(Default)'
            if (Test-Path $regPath) { $exeToStart = $regPath }
        } catch {}
    }

    if (-not $exeToStart) {
        Write-Error "Could not find Outlook (classic OUTLOOK.EXE or new olk.exe). Please open Outlook manually and retry."
        exit 1
    }

    $isClassicExe = $exeToStart -notmatch "olk\.exe$"
    Start-Process $exeToStart

    # Wait up to 60 seconds for Outlook COM to become available
    $app = $null
    for ($i = 0; $i -lt 60; $i++) {
        Start-Sleep 1
        try {
            $app = if ($isClassicExe) {
                [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            } else {
                New-Object -ComObject Outlook.Application
            }
            Write-Host "[export-calendar] Outlook started and ready after $($i+1)s."
            break
        } catch {
            # Still starting
        }
    }

    if (-not $app) {
        Write-Error "Outlook did not become ready within 60 seconds."
        exit 1
    }

    # Give Outlook a moment to fully load the profile/mailbox
    Start-Sleep 3
    return $app
}

# ---------------------------------------------------------------------------
# Parse and validate dates
# ---------------------------------------------------------------------------
try {
    $start = [DateTime]::Parse($StartDate)
} catch {
    Write-Error "Invalid StartDate: '$StartDate'. Use YYYY-MM-DD format."
    exit 1
}

try {
    # EndDate is inclusive - set to end of day
    $end = [DateTime]::Parse($EndDate).Date.AddDays(1).AddSeconds(-1)
} catch {
    Write-Error "Invalid EndDate: '$EndDate'. Use YYYY-MM-DD format."
    exit 1
}

Write-Host "[export-calendar] Date range: $($start.ToString('yyyy-MM-dd')) to $([DateTime]::Parse($EndDate).ToString('yyyy-MM-dd'))"
Write-Host "[export-calendar] Skip private: $SkipPrivate"

# ---------------------------------------------------------------------------
# Ensure output directory exists
# ---------------------------------------------------------------------------
$outDir = Split-Path -Parent $OutFile
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
    Write-Host "[export-calendar] Created directory: $outDir"
}

# ---------------------------------------------------------------------------
# Connect to Outlook and export
# ---------------------------------------------------------------------------
$outlook = Get-OutlookCOM
$ns = $outlook.GetNamespace("MAPI")

# olFolderCalendar = 9
$calendar = $ns.GetDefaultFolder(9)

$exporter = $calendar.GetCalendarExporter()

# olFullDetails = 2 (includes subject, location, body, attendees)
$exporter.CalendarDetail = 2
$exporter.IncludeAttachments = $false
$exporter.IncludePrivateDetails = (-not $SkipPrivate)
$exporter.IncludeWholeCalendar = $false
$exporter.RestrictToWorkingHours = $false
$exporter.StartDate = $start
$exporter.EndDate = $end

try {
    $exporter.SaveAsICal($OutFile)
    Write-Host "[export-calendar] Exported successfully to: $OutFile"
} catch {
    Write-Error "Export failed: $_"
    exit 1
} finally {
    # Release COM objects to avoid memory leaks
    if ($exporter) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($exporter) | Out-Null }
    if ($calendar) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($calendar) | Out-Null }
    if ($ns)       { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns) | Out-Null }
    # Do NOT release $outlook - we don't want to close a running Outlook instance
}

# Verify file was created
if (Test-Path $OutFile) {
    $size = (Get-Item $OutFile).Length
    Write-Host "[export-calendar] File size: $size bytes"
    exit 0
} else {
    Write-Error "Export appeared to succeed but file not found at: $OutFile"
    exit 1
}
