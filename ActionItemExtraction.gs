/**
 * Meeting action-item extraction pipeline (see README's Roadmap section).
 *
 * Piece 1/4: poll recent Calendar events, wait for the Gemini notes doc to
 * show up as an attachment, then pull its text. Tracked in a new
 * `ProcessedMeetings` tab so a meeting is never re-scanned once its notes
 * doc has been found and read.
 *
 * Requires the advanced "Google Calendar API" service (Services > + >
 * Google Calendar API) - same as the spike script.
 */

const ACTION_ITEM_SHEET_NAMES = {
  PROCESSED_MEETINGS: "ProcessedMeetings",
};

const MEETING_LOOKBACK_HOURS = 24;

function ensureActionItemSheets_() {
  const ss = SpreadsheetApp.openById(getSheetId_());
  ensureSheet_(ss, ACTION_ITEM_SHEET_NAMES.PROCESSED_MEETINGS, [
    "EventID", "Title", "Status", "FirstSeenAt", "LastCheckedAt", "NotesDocId",
  ]);
}

/**
 * Run once manually to create the ProcessedMeetings tab, then install the
 * time-driven trigger.
 */
function setupActionItemExtraction() {
  ensureActionItemSheets_();
  Logger.log("ProcessedMeetings tab ready.");
}

/**
 * Installs (or reinstalls) the 15-minute polling trigger.
 */
function installMeetingPollTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "pollMeetingNotes") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("pollMeetingNotes").timeBased().everyMinutes(15).create();
  Logger.log("Installed 15-minute pollMeetingNotes trigger.");
}

/**
 * Time-driven entry point. Sweeps events from the last MEETING_LOOKBACK_HOURS
 * hours; for any not yet marked "notes_fetched", checks for a Gemini notes
 * doc attachment and pulls its text once found.
 */
function pollMeetingNotes() {
  ensureActionItemSheets_();
  const sheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PROCESSED_MEETINGS);
  const tracked = rowsToObjects_(sheet);
  const trackedByEventId = {};
  tracked.forEach(function (row) {
    trackedByEventId[row.EventID] = row;
  });

  const now = new Date();
  const lookbackStart = new Date(now.getTime() - MEETING_LOOKBACK_HOURS * 60 * 60 * 1000);

  const events = Calendar.Events.list("primary", {
    timeMin: lookbackStart.toISOString(),
    timeMax: now.toISOString(),
    singleEvents: true,
    orderBy: "startTime",
    maxResults: 100,
  });

  const items = events.items || [];
  var foundCount = 0;

  items.forEach(function (ev) {
    const existing = trackedByEventId[ev.id];
    if (existing && existing.Status === "notes_fetched") return;

    const notesDoc = findNotesDocAttachment_(ev);

    if (!notesDoc) {
      upsertProcessedMeeting_(sheet, existing, ev, "waiting", "");
      return;
    }

    upsertProcessedMeeting_(sheet, existing, ev, "notes_fetched", notesDoc.fileId);
    foundCount++;
    Logger.log('Notes doc found for "%s" (event %s): fileId=%s', ev.summary, ev.id, notesDoc.fileId);
  });

  Logger.log("Poll complete. %s event(s) newly notes_fetched.", foundCount);
}

/**
 * Looks for a Google Doc attachment on the event that looks like meeting
 * notes. Gemini's notes doc is a Google Doc attachment; we don't rely on
 * its title matching anything specific since that's not guaranteed stable.
 */
function findNotesDocAttachment_(ev) {
  const attachments = ev.attachments || [];
  for (var i = 0; i < attachments.length; i++) {
    const att = attachments[i];
    if (att.mimeType === "application/vnd.google-apps.document" && att.fileId) {
      return att;
    }
  }
  return null;
}

function upsertProcessedMeeting_(sheet, existing, ev, status, notesDocId) {
  const now = new Date().getTime();
  if (existing) {
    updateCell_(sheet, "EventID", ev.id, "Status", status);
    updateCell_(sheet, "EventID", ev.id, "LastCheckedAt", now);
    if (notesDocId) updateCell_(sheet, "EventID", ev.id, "NotesDocId", notesDocId);
  } else {
    sheet.appendRow([ev.id, ev.summary || "", status, now, now, notesDocId || ""]);
  }
}

/**
 * Fetches the plain text of a Gemini notes doc by its Drive file ID.
 */
function getNotesDocText_(fileId) {
  return DocumentApp.openById(fileId).getBody().getText();
}

/**
 * Manual spot-check: logs the notes text length for every event currently
 * marked notes_fetched, so you can confirm real content is coming through
 * before wiring up the NIM call.
 */
function debugPrintFetchedNotes() {
  const sheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PROCESSED_MEETINGS);
  const rows = rowsToObjects_(sheet).filter(function (r) {
    return r.Status === "notes_fetched";
  });
  Logger.log("%s meeting(s) with notes fetched.", rows.length);
  rows.forEach(function (r) {
    const text = getNotesDocText_(r.NotesDocId);
    Logger.log('--- "%s" (%s chars) ---\n%s', r.Title, text.length, text.substring(0, 500));
  });
}
