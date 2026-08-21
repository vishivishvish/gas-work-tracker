/**
 * Meeting action-item extraction pipeline (see README's Roadmap section).
 *
 * Piece 1/4: poll recent Calendar events, wait for the Gemini notes doc to
 * show up as an attachment, then pull its text. Tracked in a new
 * `ProcessedMeetings` tab so a meeting is never re-scanned once its notes
 * doc has been found and read.
 *
 * Piece 2/4: send each meeting's notes text (plus the board's current
 * threads/sub-threads) to NVIDIA NIM, which proposes action items with a
 * suggested thread/sub-thread placement and a rationale.
 *
 * Piece 3/4: expand those proposals into a `PendingActions` tab (one row per
 * item, each with its own review token) and email a review link per item.
 * Nothing lands on the board yet - see piece 4 for the two-step approval
 * that actually writes to Threads/SubThreads/Items.
 *
 * Requires the advanced "Google Calendar API" service (Services > + >
 * Google Calendar API) - same as the spike script.
 */

const ACTION_ITEM_SHEET_NAMES = {
  PROCESSED_MEETINGS: "ProcessedMeetings",
  PENDING_ACTIONS: "PendingActions",
};

const MEETING_LOOKBACK_HOURS = 24;

// NIM's OpenAI-compatible chat completions endpoint. Model is a Script
// Property (NIM_MODEL) rather than hardcoded, since NIM's catalog of hosted
// Nemotron/other models changes - set it to whatever model ID you've picked
// on build.nvidia.com. NIM_API_KEY is required; get one from build.nvidia.com.
const NIM_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions";

function ensureActionItemSheets_() {
  const ss = SpreadsheetApp.openById(getSheetId_());
  ensureSheet_(ss, ACTION_ITEM_SHEET_NAMES.PROCESSED_MEETINGS, [
    "EventID", "Title", "Status", "FirstSeenAt", "LastCheckedAt", "NotesDocId", "ProposalsJson",
  ]);
  ensureColumn_(getSheet_(ACTION_ITEM_SHEET_NAMES.PROCESSED_MEETINGS), "ProposalsJson");

  ensureSheet_(ss, ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS, [
    "Token", "EventID", "MeetingTitle", "Text", "Owner",
    "ThreadName", "IsNewThread", "SubThreadName", "IsNewSubThread", "SubThreadTag",
    "Rationale", "Status", "CreatedAt",
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
  ensureActionItemSheets_();
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

// ---- Piece 2/4: NIM call - propose action items + thread/sub-thread match ----

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
 * parsed array of proposed action items. Throws on any failure (HTTP error,
 * missing API key, unparsable response) so callers can leave the meeting's
 * Status untouched and retry on the next poll.
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
 * Runs NIM extraction for every meeting currently marked notes_fetched,
 * storing the raw proposals JSON back onto its ProcessedMeetings row and
 * advancing Status to "extracted". Leaves failures at notes_fetched so the
 * next manual run (or future scheduled run) retries them.
 */
function extractActionItemsForFetchedMeetings() {
  ensureActionItemSheets_();
  const sheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PROCESSED_MEETINGS);
  const rows = rowsToObjects_(sheet).filter(function (r) {
    return r.Status === "notes_fetched";
  });

  Logger.log("%s meeting(s) to extract.", rows.length);

  rows.forEach(function (r) {
    try {
      const notesText = getNotesDocText_(r.NotesDocId);
      const items = callNimForActionItems_(r.Title, notesText);
      updateCell_(sheet, "EventID", r.EventID, "ProposalsJson", JSON.stringify(items));
      updateCell_(sheet, "EventID", r.EventID, "Status", "extracted");
      Logger.log('Extracted %s item(s) for "%s".', items.length, r.Title);
    } catch (err) {
      Logger.log('Extraction failed for "%s": %s', r.Title, err.message);
    }
  });
}

// ---- Piece 3/4: PendingActions tab + review email ----

/**
 * Expands each "extracted" meeting's ProposalsJson into one PendingActions
 * row per item (each with its own review token), then advances the meeting
 * to "rows_created". Split from email-sending so a mid-run failure never
 * leaves a meeting stuck between "some rows written" and "no rows written".
 */
function createPendingActionsForExtractedMeetings() {
  ensureActionItemSheets_();
  const meetingsSheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PROCESSED_MEETINGS);
  const pendingSheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS);
  const meetings = rowsToObjects_(meetingsSheet).filter(function (r) {
    return r.Status === "extracted";
  });

  Logger.log("%s meeting(s) to turn into pending actions.", meetings.length);

  meetings.forEach(function (meeting) {
    var items;
    try {
      items = JSON.parse(meeting.ProposalsJson);
    } catch (err) {
      Logger.log('Skipping "%s": ProposalsJson did not parse (%s).', meeting.Title, err.message);
      return;
    }

    const now = new Date().getTime();
    items.forEach(function (item) {
      pendingSheet.appendRow([
        Utilities.getUuid(),
        meeting.EventID,
        meeting.Title,
        item.text || "",
        item.owner || "",
        item.threadName || "",
        !!item.isNewThread,
        item.subThreadName || "",
        !!item.isNewSubThread,
        item.subThreadTag || "",
        item.rationale || "",
        "proposed",
        now,
      ]);
    });

    updateCell_(meetingsSheet, "EventID", meeting.EventID, "Status", "rows_created");
    Logger.log('Created %s pending action row(s) for "%s".', items.length, meeting.Title);
  });
}

/**
 * Sends one review email per "rows_created" meeting, listing its pending
 * items with a per-item review link, then advances the meeting to "emailed".
 * If WEBAPP_URL isn't set yet (the approval web app doesn't exist until
 * piece 4 is deployed), this logs a warning and leaves the meeting at
 * "rows_created" so it retries once WEBAPP_URL is set.
 */
function sendPendingActionEmails() {
  ensureActionItemSheets_();
  const props = PropertiesService.getScriptProperties();
  const webAppUrl = props.getProperty("WEBAPP_URL");
  if (!webAppUrl) {
    Logger.log("WEBAPP_URL not set yet - skipping email send until the approval web app is deployed.");
    return;
  }

  const meetingsSheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PROCESSED_MEETINGS);
  const pendingSheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS);
  const meetings = rowsToObjects_(meetingsSheet).filter(function (r) {
    return r.Status === "rows_created";
  });
  const allPending = rowsToObjects_(pendingSheet);

  Logger.log("%s meeting(s) to email.", meetings.length);

  meetings.forEach(function (meeting) {
    const items = allPending.filter(function (p) {
      return p.EventID === meeting.EventID && p.Status === "proposed";
    });
    if (!items.length) {
      updateCell_(meetingsSheet, "EventID", meeting.EventID, "Status", "emailed");
      return;
    }

    const htmlBody = buildReviewEmailHtml_(meeting.Title, items, webAppUrl);
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: "Action items to review: " + meeting.Title,
      htmlBody: htmlBody,
    });

    updateCell_(meetingsSheet, "EventID", meeting.EventID, "Status", "emailed");
    Logger.log('Emailed %s item(s) for "%s".', items.length, meeting.Title);
  });
}

function buildReviewEmailHtml_(meetingTitle, items, webAppUrl) {
  const rows = items
    .map(function (item) {
      const threadLabel = escapeHtml_(item.ThreadName) + (item.IsNewThread ? " <b>(NEW)</b>" : "");
      const subThreadLabel = escapeHtml_(item.SubThreadName) + (item.IsNewSubThread ? " <b>(NEW)</b>" : "");
      const reviewLink = webAppUrl + "?token=" + encodeURIComponent(item.Token);
      return (
        "<tr>" +
        "<td>" + escapeHtml_(item.Text) + "</td>" +
        "<td>" + escapeHtml_(item.Owner || "(unassigned)") + "</td>" +
        "<td>" + threadLabel + " &rsaquo; " + subThreadLabel + "</td>" +
        "<td><i>" + escapeHtml_(item.Rationale) + "</i></td>" +
        '<td><a href="' + reviewLink + '">Review</a></td>' +
        "</tr>"
      );
    })
    .join("");

  return (
    "<p>Proposed action items from <b>" + escapeHtml_(meetingTitle) + "</b>:</p>" +
    "<table border='1' cellpadding='6' cellspacing='0'>" +
    "<tr><th>Item</th><th>Owner</th><th>Proposed thread &rsaquo; sub-thread</th><th>Why</th><th></th></tr>" +
    rows +
    "</table>" +
    "<p>Click Review on each item to accept it as-is or edit it. " +
    "Thread/sub-thread placement isn't final yet - you'll get a separate " +
    "chance to double-check and commit that before anything hits the board.</p>"
  );
}

function escapeHtml_(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
