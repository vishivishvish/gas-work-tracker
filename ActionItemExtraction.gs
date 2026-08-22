/**
 * Meeting action-item extraction (see README's Roadmap section). Fully
 * manual now - no automatic polling/trigger, no email. You browse your
 * Calendar in the "Meetings" tab, click "Extract Action Items" on a meeting
 * that has a Gemini notes attachment, review/edit the results in the
 * "Action Items Pending for Approval" tab, and Commit or Dismiss each one.
 *
 * A meeting can only ever be extracted once - `ExtractedMeetings` tracks
 * that permanently by EventID, independent of what happens to its
 * PendingActions rows afterward (committed, dismissed, whatever), so
 * there's no automatic process left that could re-trigger it and no way to
 * click the button twice for the same meeting.
 *
 * Requires the advanced "Google Calendar API" service (Services > + >
 * Google Calendar API) - same as the spike script.
 */

const ACTION_ITEM_SHEET_NAMES = {
  EXTRACTED_MEETINGS: "ExtractedMeetings",
  PENDING_ACTIONS: "PendingActions",
};

// The day you joined Great Learning - meetings before this are never listed.
const MEETINGS_SINCE_DATE = "2026-05-19T00:00:00";

// NIM's OpenAI-compatible chat completions endpoint. Model is a Script
// Property (NIM_MODEL) rather than hardcoded, since NIM's catalog of hosted
// Nemotron/other models changes - set it to whatever model ID you've picked
// on build.nvidia.com. NIM_API_KEY is required; get one from build.nvidia.com.
const NIM_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions";

function ensureActionItemSheets_() {
  const ss = SpreadsheetApp.openById(getSheetId_());
  ensureSheet_(ss, ACTION_ITEM_SHEET_NAMES.EXTRACTED_MEETINGS, ["EventID", "Title", "ExtractedAt"]);
  ensureSheet_(ss, ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS, [
    "Token", "EventID", "MeetingTitle", "Text", "Owner", "OpenDate",
    "ThreadName", "IsNewThread", "SubThreadName", "IsNewSubThread", "SubThreadTag",
    "Rationale", "Status", "CreatedAt",
  ]);
}

/**
 * One-off reset: run this to completely clear out this feature's state -
 * deletes the old ProcessedMeetings tab entirely (belonged to the retired
 * automatic-polling design), clears PendingActions and ExtractedMeetings
 * down to just their headers, and removes any leftover trigger from the old
 * design. Does NOT touch the Work Tracker board itself (Threads/SubThreads/
 * Items/Meta) - anything already committed to the board stays exactly as is.
 */
function wipeActionItemPipeline() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === "runActionItemPipeline" || fn === "pollMeetingNotes") {
      ScriptApp.deleteTrigger(t);
    }
  });

  const ss = SpreadsheetApp.openById(getSheetId_());
  const oldSheet = ss.getSheetByName("ProcessedMeetings");
  if (oldSheet) ss.deleteSheet(oldSheet);

  [ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS, ACTION_ITEM_SHEET_NAMES.EXTRACTED_MEETINGS].forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
  });

  ensureActionItemSheets_();
  Logger.log("Wiped: removed old ProcessedMeetings tab, cleared PendingActions and ExtractedMeetings, removed any leftover trigger. Work Tracker board untouched.");
}

/**
 * Appends a row built from a {headerName: value} map rather than a fixed
 * positional array, reading the sheet's actual current header order. Safer
 * than a raw appendRow([...]) once ensureColumn_ has bolted extra columns
 * onto an existing sheet in whatever order they were added.
 */
function appendRowByHeaders_(sheet, valuesByHeader) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const row = headers.map(function (h) {
    return valuesByHeader.hasOwnProperty(h) ? valuesByHeader[h] : "";
  });
  sheet.appendRow(row);
}

/**
 * Runs fn() while holding the script-wide lock and returns its result, so
 * two overlapping calls (e.g. an accidental double-click on "Extract Action
 * Items") can't both act on the same meeting. Returns null without calling
 * fn() if the lock is already held.
 */
function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("Another run is already in progress - skipping this run.");
    return null;
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ---- Calendar browsing (the "Meetings" tab) ----

/**
 * An attachment is treated as a Gemini transcript if its title starts with
 * "Notes" and contains "Gemini" - the naming convention Gemini's own
 * meeting notes docs use.
 */
function looksLikeGeminiTranscript_(title) {
  const t = String(title || "").trim();
  return /^notes/i.test(t) && /gemini/i.test(t);
}

function getEventColorMap_() {
  try {
    return Calendar.Colors.get().event || {};
  } catch (err) {
    return {};
  }
}

function getPrimaryCalendarColor_() {
  try {
    const cal = Calendar.CalendarList.get("primary");
    return { background: cal.backgroundColor, foreground: cal.foregroundColor };
  } catch (err) {
    return { background: "#4285F4", foreground: "#ffffff" };
  }
}

function getExtractedEventIdSet_() {
  const set = {};
  rowsToObjects_(getSheet_(ACTION_ITEM_SHEET_NAMES.EXTRACTED_MEETINGS)).forEach(function (r) {
    set[r.EventID] = true;
  });
  return set;
}

/**
 * Lists every Calendar event since MEETINGS_SINCE_DATE (never later than
 * right now, regardless of what's still ahead on the calendar) that has at
 * least one attachment, reverse-chronological. Each meeting's attachments
 * are flagged for whether they look like a Gemini transcript, and the
 * meeting is flagged with whether "Extract Action Items" has already been
 * run for it (so the tab can gray out that button).
 */
function getCalendarMeetingsWithAttachments() {
  ensureActionItemSheets_();

  const timeMin = new Date(MEETINGS_SINCE_DATE).toISOString();
  const timeMax = new Date().toISOString();
  const colorMap = getEventColorMap_();
  const defaultColor = getPrimaryCalendarColor_();
  const extractedIds = getExtractedEventIdSet_();

  var events = [];
  var pageToken;
  do {
    const resp = Calendar.Events.list("primary", {
      timeMin: timeMin,
      timeMax: timeMax,
      singleEvents: true,
      orderBy: "startTime",
      maxResults: 250,
      pageToken: pageToken,
    });
    events = events.concat(resp.items || []);
    pageToken = resp.nextPageToken;
  } while (pageToken);

  const meetings = events
    .filter(function (ev) { return (ev.attachments || []).length > 0; })
    .map(function (ev) {
      const attachments = (ev.attachments || []).map(function (a) {
        return {
          title: a.title || "(untitled attachment)",
          mimeType: a.mimeType || "",
          looksLikeTranscript: looksLikeGeminiTranscript_(a.title),
        };
      });
      const hasTranscript = attachments.some(function (a) { return a.looksLikeTranscript; });
      const color = (ev.colorId && colorMap[ev.colorId]) ? colorMap[ev.colorId] : defaultColor;

      return {
        eventId: ev.id,
        title: ev.summary || "(untitled meeting)",
        start: ev.start ? (ev.start.dateTime || ev.start.date) : "",
        end: ev.end ? (ev.end.dateTime || ev.end.date) : "",
        isAllDay: !!(ev.start && ev.start.date && !ev.start.dateTime),
        attendees: (ev.attendees || []).map(function (a) { return a.displayName || a.email; }),
        attachments: attachments,
        hasTranscript: hasTranscript,
        colorBackground: color.background,
        colorForeground: color.foreground,
        extracted: !!extractedIds[ev.id],
      };
    });

  meetings.sort(function (a, b) { return new Date(b.start) - new Date(a.start); });
  return meetings;
}

/**
 * Picks the attachment to use as the meeting's notes doc: prefers one
 * matching the Gemini-transcript naming convention, falling back to any
 * Google Doc attachment if that convention isn't matched by anything.
 */
function findNotesDocAttachment_(ev) {
  const attachments = ev.attachments || [];
  const byConvention = attachments.filter(function (a) {
    return a.mimeType === "application/vnd.google-apps.document" && looksLikeGeminiTranscript_(a.title);
  })[0];
  if (byConvention) return byConvention;
  return attachments.filter(function (a) { return a.mimeType === "application/vnd.google-apps.document"; })[0] || null;
}

/**
 * The notes doc's own creation date (YYYY-MM-DD), used as every action
 * item's default opening date.
 */
function getNotesDocDate_(fileId) {
  const created = DriveApp.getFileById(fileId).getDateCreated();
  return Utilities.formatDate(created, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function getNotesDocText_(fileId) {
  return DocumentApp.openById(fileId).getBody().getText();
}

// ---- NIM call - propose action items + thread/sub-thread match ----

/**
 * Compact text summary of the current board's threads/sub-threads, given to
 * the LLM so it matches against real structure instead of guessing blind.
 */
function buildBoardContextText_() {
  const board = getBoardData().board;
  if (!board.length) return "(no threads exist yet)";
  return board
    .map(function (t) {
      const subLines = t.subthreads
        .map(function (s) {
          return "  - " + s.name + (s.tag ? " (tag: " + s.tag + ")" : "");
        })
        .join("\n");
      return "- " + t.name + (subLines ? "\n" + subLines : "");
    })
    .join("\n");
}

function buildExtractionPrompt_(meetingTitle, notesText, boardContextText) {
  return (
    "You are extracting action items from a meeting's notes for a work " +
    "tracker. The tracker organizes work as Threads, each containing " +
    "Sub-Threads, each containing Action Items.\n\n" +
    "Current threads and sub-threads:\n" + boardContextText + "\n\n" +
    'Meeting title: "' + meetingTitle + '"\n\n' +
    "Meeting notes:\n" + notesText + "\n\n" +
    "Extract every concrete action item from these notes. For each one, " +
    "decide which existing thread and sub-thread it belongs to. Only " +
    "propose a new thread or sub-thread if none of the existing ones are a " +
    "reasonable fit - prefer reusing existing structure. " +
    "Respond with ONLY a JSON array (no markdown fences, no commentary), " +
    "where each element has exactly these fields:\n" +
    '{"text": string, "owner": string (best-guess name, or "" if unclear), ' +
    '"threadName": string, "isNewThread": boolean, ' +
    '"subThreadName": string, "isNewSubThread": boolean, ' +
    '"subThreadTag": string (short tag, only meaningful if isNewSubThread), ' +
    '"rationale": string (one short sentence on why this thread/sub-thread ' +
    "was chosen, to help a human sanity-check the placement)}\n" +
    "If there are no action items, respond with an empty JSON array: []"
  );
}

/**
 * Calls NIM's OpenAI-compatible chat completions endpoint and returns the
 * parsed array of proposed action items. Throws on any failure.
 */
function callNimForActionItems_(meetingTitle, notesText) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("NIM_API_KEY");
  const model = props.getProperty("NIM_MODEL");
  if (!apiKey) throw new Error("NIM_API_KEY not set in Script Properties.");
  if (!model) throw new Error("NIM_MODEL not set in Script Properties.");

  const prompt = buildExtractionPrompt_(meetingTitle, notesText, buildBoardContextText_());

  const response = UrlFetchApp.fetch(NIM_API_URL, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      model: model,
      messages: [{ role: "user", content: prompt }],
      temperature: 0.2,
    }),
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("NIM API error " + code + ": " + response.getContentText());
  }

  const parsed = JSON.parse(response.getContentText());
  const content = parsed.choices && parsed.choices[0] && parsed.choices[0].message
    ? parsed.choices[0].message.content
    : null;
  if (!content) throw new Error("NIM response missing choices[0].message.content.");

  return parseActionItemsJson_(content);
}

/**
 * NIM/Nemotron models sometimes wrap JSON in markdown fences despite being
 * told not to - strip those before parsing.
 */
function parseActionItemsJson_(content) {
  const cleaned = content.trim().replace(/^```(?:json)?/i, "").replace(/```$/, "").trim();
  const items = JSON.parse(cleaned);
  if (!Array.isArray(items)) throw new Error("NIM response was not a JSON array.");
  return items;
}

/**
 * Called from the Meetings tab's "Extract Action Items" button. Runs NIM
 * extraction for this one meeting and writes its proposals into
 * PendingActions, then permanently marks the meeting in ExtractedMeetings -
 * the Meetings tab checks that to gray the button out and refuses to let it
 * be clicked again, which is the only guard needed now that there's no
 * automatic process left that could re-trigger this on its own.
 */
function extractActionItemsForMeeting(eventId) {
  return withScriptLock_(function () {
    ensureActionItemSheets_();

    if (getExtractedEventIdSet_()[eventId]) {
      throw new Error("Action items have already been extracted for this meeting.");
    }

    const ev = Calendar.Events.get("primary", eventId);
    const notesDoc = findNotesDocAttachment_(ev);
    if (!notesDoc) throw new Error("No Gemini notes attachment found on this meeting.");

    const notesText = getNotesDocText_(notesDoc.fileId);
    const notesDate = getNotesDocDate_(notesDoc.fileId);
    const items = callNimForActionItems_(ev.summary || "", notesText);

    const pendingSheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS);
    const now = new Date().getTime();
    items.forEach(function (item) {
      appendRowByHeaders_(pendingSheet, {
        Token: Utilities.getUuid(),
        EventID: eventId,
        MeetingTitle: ev.summary || "",
        Text: item.text || "",
        Owner: item.owner || "",
        OpenDate: notesDate || "",
        ThreadName: item.threadName || "",
        IsNewThread: !!item.isNewThread,
        SubThreadName: item.subThreadName || "",
        IsNewSubThread: !!item.isNewSubThread,
        SubThreadTag: item.subThreadTag || "",
        Rationale: item.rationale || "",
        Status: "proposed",
        CreatedAt: now,
      });
    });

    appendRowByHeaders_(getSheet_(ACTION_ITEM_SHEET_NAMES.EXTRACTED_MEETINGS), {
      EventID: eventId,
      Title: ev.summary || "",
      ExtractedAt: now,
    });

    Logger.log('Extracted %s item(s) for "%s".', items.length, ev.summary);
    return items.length;
  });
}

// ---- In-app proposal review (the "Action Items Pending for Approval" tab) ----

function getPendingActionRow_(token) {
  const rows = rowsToObjects_(getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS));
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].Token === token) return rows[i];
  }
  return null;
}

/**
 * Returns every "proposed" item for the pending-approval tab to render. The
 * tab itself disappears from the tab bar whenever this comes back empty.
 */
function getPendingProposals() {
  return rowsToObjects_(getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS)).filter(function (r) {
    return r.Status === "proposed";
  });
}

/**
 * Called from the tab's Reject button - discards a proposal without
 * touching the board.
 */
function dismissProposal(token) {
  const sheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS);
  const row = getPendingActionRow_(token);
  if (!row) throw new Error("Pending action not found.");
  if (row.Status !== "proposed") throw new Error("This item was already reviewed.");
  updateCell_(sheet, "Token", token, "Status", "rejected");
}

/**
 * Called from the tab's Accept (with edits) button. Fields always reflect
 * the final intended values, so this just finds-or-creates the chosen
 * thread/sub-thread and writes the item straight to the board via the same
 * addThread/addSubThread/addItem functions Code.gs already uses, then marks
 * the PendingActions row committed.
 */
function submitItemDecision(token, text, owner, openDate, closeDate, threadName, isNewThread, subThreadName, isNewSubThread, subThreadTag) {
  const sheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS);
  const row = getPendingActionRow_(token);
  if (!row) throw new Error("Pending action not found.");
  if (row.Status !== "proposed") throw new Error("This item was already reviewed.");

  const threadId = findOrCreateThread_(threadName);
  const subThreadId = findOrCreateSubThread_(threadId, subThreadName, subThreadTag);
  addItem(subThreadId, text, owner, openDate, closeDate || openDate);

  updateCell_(sheet, "Token", token, "Status", "committed");
}

function findOrCreateThread_(name) {
  const trimmedName = String(name || "").trim();
  const match = rowsToObjects_(getSheet_(SHEET_NAMES.THREADS)).find(function (t) {
    return String(t.Name).trim().toLowerCase() === trimmedName.toLowerCase();
  });
  return match ? match.ThreadID : addThread(trimmedName, "");
}

function findOrCreateSubThread_(threadId, name, tag) {
  const trimmedName = String(name || "").trim();
  const match = rowsToObjects_(getSheet_(SHEET_NAMES.SUBTHREADS)).find(function (s) {
    return s.ThreadID === threadId && String(s.Name).trim().toLowerCase() === trimmedName.toLowerCase();
  });
  return match ? match.SubThreadID : addSubThread(threadId, trimmedName, tag || "", "");
}
