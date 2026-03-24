#!/usr/bin/env node
/**
 * parse-ics.js
 * Reads and parses a locally exported Outlook ICS calendar file.
 * Filters by date range and optionally skips private events.
 *
 * Usage:
 *   node parse-ics.js --range this-week [--skip-private] [--file path/to/calendar.ics]
 *   node parse-ics.js --range today [--skip-private]
 *   node parse-ics.js --range custom --start 2026-03-20 --end 2026-03-27 [--skip-private]
 *
 * The ICS file is produced by export-calendar.ps1 before calling this script.
 * Default file location: %USERPROFILE%\.mytime-booker\calendar.ics
 *
 * Output: JSON array of events to stdout
 */

const fs = require("fs");
const path = require("path");

// ---------------------------------------------------------------------------
// CLI argument parsing
// ---------------------------------------------------------------------------
const args = process.argv.slice(2);
const get = (flag) => {
  const i = args.indexOf(flag);
  return i !== -1 ? args[i + 1] : null;
};
const has = (flag) => args.includes(flag);

const range = get("--range") || "this-week";
const skipPrivate = has("--skip-private");
const customStart = get("--start");
const customEnd = get("--end");

const defaultIcsPath = path.join(
  process.env.USERPROFILE || process.env.HOME || ".",
  ".mytime-booker",
  "calendar.ics"
);
const icsPath = get("--file") || defaultIcsPath;

// ---------------------------------------------------------------------------
// Date range calculation
// ---------------------------------------------------------------------------
function getDateRange(range, customStart, customEnd) {
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  if (range === "today") {
    const end = new Date(today);
    end.setHours(23, 59, 59, 999);
    return { start: today, end };
  }

  if (range === "this-week") {
    // Week starts Monday
    const day = today.getDay();
    const diffToMonday = day === 0 ? -6 : 1 - day;
    const monday = new Date(today);
    monday.setDate(today.getDate() + diffToMonday);
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);
    sunday.setHours(23, 59, 59, 999);
    return { start: monday, end: sunday };
  }

  if (range === "custom") {
    if (!customStart || !customEnd) {
      console.error("--range custom requires --start YYYY-MM-DD and --end YYYY-MM-DD");
      process.exit(1);
    }
    const start = new Date(customStart + "T00:00:00");
    const end = new Date(customEnd + "T23:59:59");
    if (isNaN(start) || isNaN(end)) {
      console.error("Invalid date format. Use YYYY-MM-DD");
      process.exit(1);
    }
    return { start, end };
  }

  console.error(`Unknown range: ${range}. Use: today | this-week | custom`);
  process.exit(1);
}

// ---------------------------------------------------------------------------
// ICS parser
// ---------------------------------------------------------------------------
function unfoldLines(raw) {
  // ICS lines fold at 75 chars with CRLF + whitespace continuation
  return raw.replace(/\r?\n[ \t]/g, "");
}

function parseICSDate(value) {
  // Handles: 20260323T090000Z, 20260323T090000, 20260323
  if (!value) return null;
  const clean = value.replace(/^TZID=[^:]+:/, ""); // strip TZID param if present
  const isUTC = clean.endsWith("Z");
  const d = clean.replace("Z", "");

  if (d.length === 8) {
    // All-day: YYYYMMDD
    return new Date(
      parseInt(d.slice(0, 4)),
      parseInt(d.slice(4, 6)) - 1,
      parseInt(d.slice(6, 8))
    );
  }

  // Datetime: YYYYMMDDTHHmmss
  const dt = new Date(
    Date.UTC(
      parseInt(d.slice(0, 4)),
      parseInt(d.slice(4, 6)) - 1,
      parseInt(d.slice(6, 8)),
      parseInt(d.slice(9, 11)),
      parseInt(d.slice(11, 13)),
      parseInt(d.slice(13, 15))
    )
  );

  // If not explicitly UTC and not a Z-suffixed value, treat as local
  if (!isUTC) {
    const localOffset = new Date().getTimezoneOffset() * 60000;
    return new Date(dt.getTime() + localOffset);
  }

  return dt;
}

function parsePropertyLine(line) {
  // Split on first colon, handle VALUE=DATE:, TZID=...: etc.
  const colonIdx = line.indexOf(":");
  if (colonIdx === -1) return { key: line, params: {}, value: "" };

  const keyPart = line.slice(0, colonIdx);
  const value = line.slice(colonIdx + 1);

  const semicolonIdx = keyPart.indexOf(";");
  const key = semicolonIdx === -1 ? keyPart : keyPart.slice(0, semicolonIdx);
  const paramStr = semicolonIdx === -1 ? "" : keyPart.slice(semicolonIdx + 1);

  const params = {};
  if (paramStr) {
    for (const p of paramStr.split(";")) {
      const [pk, pv] = p.split("=");
      if (pk) params[pk] = pv || "";
    }
  }

  return { key, params, value };
}

function parseICS(raw) {
  const lines = unfoldLines(raw).split(/\r?\n/);
  const events = [];
  let current = null;

  for (const line of lines) {
    if (line === "BEGIN:VEVENT") {
      current = {};
      continue;
    }
    if (line === "END:VEVENT") {
      if (current) events.push(current);
      current = null;
      continue;
    }
    if (!current) continue;

    const { key, params, value } = parsePropertyLine(line);

    switch (key) {
      case "SUMMARY":
        current.title = value.replace(/\\,/g, ",").replace(/\\n/g, "\n").trim();
        break;
      case "DTSTART":
        current.dtstart = value;
        current.dtstart_params = params;
        current.start = parseICSDate(
          params.TZID ? `TZID=${params.TZID}:${value}` : value
        );
        current.allDay = !value.includes("T");
        break;
      case "DTEND":
        current.dtend = value;
        current.end = parseICSDate(
          params.TZID ? `TZID=${params.TZID}:${value}` : value
        );
        break;
      case "DURATION":
        current.duration_raw = value;
        break;
      case "CLASS":
        current.classification = value.toUpperCase(); // PUBLIC, PRIVATE, CONFIDENTIAL
        break;
      case "STATUS":
        current.status = value.toUpperCase(); // CONFIRMED, CANCELLED, TENTATIVE
        break;
      case "TRANSP":
        current.transp = value.toUpperCase(); // TRANSPARENT (free), OPAQUE (busy)
        break;
      case "LOCATION":
        current.location = value.replace(/\\,/g, ",").trim();
        break;
      case "DESCRIPTION": {
        const parsed = value.replace(/\\n/g, "\n").replace(/\\,/g, ",").trim();
        // Outlook writes two DESCRIPTION fields: real content first, then "Reminder"
        // Keep whichever is longer (the real one)
        if (!current.description || parsed.length > current.description.length) {
          current.description = parsed;
        }
        break;
      }
      case "ORGANIZER":
        current.organizer = value.replace(/^mailto:/i, "");
        break;
      case "ATTENDEE":
        if (!current.attendees) current.attendees = [];
        current.attendees.push(value.replace(/^mailto:/i, ""));
        break;
      case "UID":
        current.uid = value;
        break;
      case "RRULE":
        current.recurring = true;
        current.rrule = value;
        break;
    }
  }

  return events;
}

// ---------------------------------------------------------------------------
// Duration calculation
// ---------------------------------------------------------------------------
function durationHours(start, end) {
  if (!start || !end) return null;
  const ms = end.getTime() - start.getTime();
  return Math.round((ms / 3600000) * 100) / 100;
}

// ---------------------------------------------------------------------------
// Format output event
// ---------------------------------------------------------------------------
function formatDate(d) {
  if (!d) return null;
  return d.toISOString().slice(0, 10);
}

function formatTime(d) {
  if (!d) return null;
  return d.toTimeString().slice(0, 5); // HH:MM
}

function cleanDescription(raw) {
  if (!raw) return null;
  // Strip Microsoft Teams join URLs and safelinks - just noise for time booking
  let text = raw
    .replace(/https?:\/\/eur\d+\.safelinks[^\s>]*/g, "")  // safelinks URLs
    .replace(/https?:\/\/teams\.microsoft[^\s>]*/g, "")    // Teams URLs
    .replace(/<https?:\/\/[^>]*>/g, "")                    // angle-bracket URLs
    .replace(/Jetzt an der Besprechung teilnehmen\s*/gi, "")
    .replace(/Microsoft Teams.*?Hilfe\?/gi, "")
    .replace(/Benötigen Sie Hilfe\?/gi, "")
    .replace(/_{5,}/g, "")                                 // long underline separators
    .replace(/\n{3,}/g, "\n\n")                            // collapse excessive newlines
    .trim();
  return text.length > 0 ? text : null;
}

function formatEvent(ev) {
  return {
    uid: ev.uid || null,
    title: ev.title || "(no title)",
    date: formatDate(ev.start),
    start: ev.allDay ? null : formatTime(ev.start),
    end: ev.allDay ? null : formatTime(ev.end),
    duration_hours: ev.allDay ? null : durationHours(ev.start, ev.end),
    all_day: ev.allDay || false,
    is_private: ev.classification === "PRIVATE" || ev.classification === "CONFIDENTIAL",
    status: ev.status || "CONFIRMED",
    location: ev.location || null,
    description: cleanDescription(ev.description),
    organizer: ev.organizer || null,
    attendees: ev.attendees || [],
    recurring: ev.recurring || false,
  };
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
function main() {
  const { start, end } = getDateRange(range, customStart, customEnd);

  // Read local ICS file (produced by export-calendar.ps1)
  if (!fs.existsSync(icsPath)) {
    console.error(`ICS file not found at: ${icsPath}`);
    console.error(`Run export-calendar.ps1 first to generate it.`);
    process.exit(1);
  }

  const raw = fs.readFileSync(icsPath, "utf8");
  process.stderr.write(`[parse-ics] Reading from: ${icsPath}\n`);

  // Parse
  const allEvents = parseICS(raw);

  // Deduplicate: Outlook COM export writes both a master entry and a local copy
  // for each event. Keep the one with the most complete data (has title + organizer).
  const seen = new Map();
  for (const ev of allEvents) {
    const key = ev.uid || `${ev.dtstart}-${ev.title}`;
    if (!seen.has(key)) {
      seen.set(key, ev);
    } else {
      const existing = seen.get(key);
      // Prefer the entry with more data (title, organizer, attendees)
      const descScore = (d) => (!d || d === "Reminder") ? 0 : d.length;
      const existingScore = (existing.title && existing.title !== "(no title)" ? 2 : 0)
        + (existing.organizer ? 1 : 0) + (existing.attendees?.length || 0)
        + (descScore(existing.description) > 0 ? 10 : 0);
      const newScore = (ev.title && ev.title !== "(no title)" ? 2 : 0)
        + (ev.organizer ? 1 : 0) + (ev.attendees?.length || 0)
        + (descScore(ev.description) > 0 ? 10 : 0);
      if (newScore > existingScore) seen.set(key, ev);
    }
  }
  const deduped = Array.from(seen.values());

  // Filter
  const filtered = deduped
    .filter((ev) => {
      if (!ev.start) return false;

      // Skip cancelled events (title starts with "Canceled:" or status is CANCELLED)
      if (ev.status === "CANCELLED") return false;
      if (ev.title && ev.title.startsWith("Canceled:")) return false;

      // Skip private if requested (check raw classification before formatEvent)
      if (skipPrivate && (ev.classification === "PRIVATE" || ev.classification === "CONFIDENTIAL")) return false;

      // Skip entries with no title (incomplete duplicate entries)
      if (!ev.title || ev.title === "(no title)") return false;

      // Date range filter
      const evStart = ev.start;
      return evStart >= start && evStart <= end;
    })
    .sort((a, b) => a.start - b.start)
    .map(formatEvent);

  // Output
  console.log(JSON.stringify(filtered, null, 2));

  // Summary to stderr (visible to agent but not part of JSON output)
  process.stderr.write(
    `\n[parse-ics] ${filtered.length} event(s) found between ${formatDate(start)} and ${formatDate(end)}\n`
  );
}

try {
  main();
} catch (e) {
  console.error(`Unexpected error: ${e.message}`);
  process.exit(1);
}
