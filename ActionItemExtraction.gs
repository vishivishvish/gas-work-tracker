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
 *
 * Piece 4/4: the approval web app, routed through Code.gs's doGet. A
 * `?token=...` link (from the email) opens a page to accept-as-is or edit
 * everything about the item - text, owner, opening/closing date, and the
 * thread/sub-thread it lands in - and submitting commits it straight to the
 * board via the same addThread/addSubThread/addItem functions Code.gs
 * already uses, then redirects to the Work Tracker itself.
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
    "EventID", "Title", "Status", "FirstSeenAt", "LastCheckedAt", "NotesDocId", "ProposalsJson", "NotesDate",
  ]);
  ensureColumn_(getSheet_(ACTION_ITEM_SHEET_NAMES.PROCESSED_MEETINGS), "ProposalsJson");
  ensureColumn_(getSheet_(ACTION_ITEM_SHEET_NAMES.PROCESSED_MEETINGS), "NotesDate");

  ensureSheet_(ss, ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS, [
    "Token", "EventID", "MeetingTitle", "Text", "Owner", "OpenDate",
    "ThreadName", "IsNewThread", "SubThreadName", "IsNewSubThread", "SubThreadTag",
    "Rationale", "Status", "CreatedAt",
  ]);
  ensureColumn_(getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS), "OpenDate");
}

/**
 * Appends a row built from a {headerName: value} map rather than a fixed
 * positional array, reading the sheet's actual current header order. Safer
 * than a raw appendRow([...]) once ensureColumn_ has bolted extra columns
 * onto an existing sheet in whatever order they were added, which won't
 * necessarily match a freshly-created sheet's declared header list.
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
 * Returns the URL of the currently-active web app deployment, or null if
 * this project has never been deployed as one yet. Used both for the
 * post-commit redirect on the review page and for the review email's links,
 * so there's no separate "WEBAPP_URL" Script Property to keep in sync.
 */
function getWebAppBaseUrl_() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (err) {
    return null;
  }
}

/**
 * Runs fn() while holding the script-wide lock, so two overlapping
 * executions of the same pipeline stage (the 15-minute trigger firing while
 * a manual test run is still in flight, or two trigger firings overlapping)
 * can't both read a row's status as not-yet-handled and both act on it -
 * which is what caused duplicate review emails. If the lock is already held
 * (another run is genuinely in progress), this skips fn() entirely rather
 * than waiting and re-doing work a moment later.
 */
function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("Another run is already in progress - skipping this run.");
    return;
  }
  try {
    fn();
  } finally {
    lock.releaseLock();
  }
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
  withScriptLock_(function () {
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
        upsertProcessedMeeting_(sheet, existing, ev, "waiting", "", "");
        return;
      }

      const notesDate = getNotesDocDate_(notesDoc.fileId);
      upsertProcessedMeeting_(sheet, existing, ev, "notes_fetched", notesDoc.fileId, notesDate);
      foundCount++;
      Logger.log('Notes doc found for "%s" (event %s): fileId=%s, date=%s', ev.summary, ev.id, notesDoc.fileId, notesDate);
    });

    Logger.log("Poll complete. %s event(s) newly notes_fetched.", foundCount);
  });
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

/**
 * The notes doc's own creation date (YYYY-MM-DD), used as every action
 * item's default opening date.
 */
function getNotesDocDate_(fileId) {
  const created = DriveApp.getFileById(fileId).getDateCreated();
  return Utilities.formatDate(created, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function upsertProcessedMeeting_(sheet, existing, ev, status, notesDocId, notesDate) {
  const now = new Date().getTime();
  if (existing) {
    updateCell_(sheet, "EventID", ev.id, "Status", status);
    updateCell_(sheet, "EventID", ev.id, "LastCheckedAt", now);
    if (notesDocId) updateCell_(sheet, "EventID", ev.id, "NotesDocId", notesDocId);
    if (notesDate) updateCell_(sheet, "EventID", ev.id, "NotesDate", notesDate);
  } else {
    appendRowByHeaders_(sheet, {
      EventID: ev.id,
      Title: ev.summary || "",
      Status: status,
      FirstSeenAt: now,
      LastCheckedAt: now,
      NotesDocId: notesDocId || "",
      NotesDate: notesDate || "",
    });
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
  withScriptLock_(function () {
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
  withScriptLock_(function () {
    ensureActionItemSheets_();
    const meetingsSheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PROCESSED_MEETINGS);
    const pendingSheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS);
    const meetings = rowsToObjects_(meetingsSheet).filter(function (r) {
      return r.Status === "extracted";
    });

    // Belt-and-suspenders on top of the lock: a meeting's EventID is its
    // unique token, so if PendingActions already has ANY row for it - no
    // matter what ProcessedMeetings.Status currently says - skip creating
    // more. This keeps things idempotent even if status ever gets out of
    // sync (a manual edit, a future bug, anything), not just against the
    // exact race the lock covers.
    const eventIdsWithPendingRows = {};
    rowsToObjects_(pendingSheet).forEach(function (p) {
      eventIdsWithPendingRows[p.EventID] = true;
    });

    Logger.log("%s meeting(s) to turn into pending actions.", meetings.length);

    meetings.forEach(function (meeting) {
      if (eventIdsWithPendingRows[meeting.EventID]) {
        Logger.log('Skipping "%s" - PendingActions rows already exist for this meeting (EventID %s).', meeting.Title, meeting.EventID);
        updateCell_(meetingsSheet, "EventID", meeting.EventID, "Status", "rows_created");
        return;
      }

      var items;
      try {
        items = JSON.parse(meeting.ProposalsJson);
      } catch (err) {
        Logger.log('Skipping "%s": ProposalsJson did not parse (%s).', meeting.Title, err.message);
        return;
      }

      const now = new Date().getTime();
      items.forEach(function (item) {
        appendRowByHeaders_(pendingSheet, {
          Token: Utilities.getUuid(),
          EventID: meeting.EventID,
          MeetingTitle: meeting.Title,
          Text: item.text || "",
          Owner: item.owner || "",
          OpenDate: meeting.NotesDate || "",
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

      updateCell_(meetingsSheet, "EventID", meeting.EventID, "Status", "rows_created");
      Logger.log('Created %s pending action row(s) for "%s".', items.length, meeting.Title);
    });
  });
}

/**
 * Sends one review email per "rows_created" meeting, listing its pending
 * items with a per-item review link, then advances the meeting to "emailed".
 * If this project hasn't been deployed as a web app yet (no URL to link to),
 * this logs a warning and leaves the meeting at "rows_created" so it retries
 * once it has been.
 */
function sendPendingActionEmails() {
  withScriptLock_(function () {
    ensureActionItemSheets_();
    const webAppUrl = getWebAppBaseUrl_();
    if (!webAppUrl) {
      Logger.log("No web app deployment found yet - skipping email send until this project is deployed.");
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

      try {
        const htmlBody = buildReviewEmailHtml_(meeting.Title, items, webAppUrl);
        MailApp.sendEmail({
          to: Session.getActiveUser().getEmail(),
          subject: "Action items to review: " + meeting.Title,
          htmlBody: htmlBody,
        });

        updateCell_(meetingsSheet, "EventID", meeting.EventID, "Status", "emailed");
        Logger.log('Emailed %s item(s) for "%s".', items.length, meeting.Title);
      } catch (err) {
        Logger.log('Failed to email/mark "%s" as emailed: %s. It will retry next run - check your inbox before re-running manually, in case the send itself succeeded but the status write failed.', meeting.Title, err.message);
      }
    });
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
        "<td>" + escapeHtml_(item.OpenDate) + "</td>" +
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
    "<tr><th>Item</th><th>Owner</th><th>Date</th><th>Proposed thread &rsaquo; sub-thread</th><th>Why</th><th></th></tr>" +
    rows +
    "</table>" +
    "<p>Click Review on each item to accept it as-is or edit anything about it - " +
    "text, owner, dates, or thread/sub-thread placement. Submitting there commits " +
    "it straight to the Work Tracker.</p>"
  );
}

function escapeHtml_(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ---- Piece 4/4: two-step approval web app ----
// Step A (per item, via email link): accept as-is or edit text/owner.
//   Only updates the PendingActions row (Status -> "approved"); the board
//   is untouched.
// Step B (Review & Commit screen): review/override thread & sub-thread
//   placement for every "approved" item, flagging anything that would
//   create a new thread/sub-thread, then write to the board on Commit.

function getPendingActionRow_(token) {
  const rows = rowsToObjects_(getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS));
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].Token === token) return rows[i];
  }
  return null;
}

// Same design tokens/fonts as Index.html's :root, so these pages read as
// part of the same app rather than a bare utility form.
const REVIEW_PAGE_CSS_ =
  ":root{" +
  "--paper:#071B45;--paper-raised:#EAF2FC;--paper-recessed:#D9E6F8;" +
  "--ink:#0B2A5B;--ink-faint:#6683AD;--canvas-ink:#EAF2FC;" +
  "--canvas-ink-faint:rgba(234,242,252,0.6);--amber:#B8722E;" +
  "--amber-soft:rgba(184,114,46,0.16);--line:rgba(11,42,91,0.14);" +
  "--done:#2FA36B;--danger:#B0554A;" +
  "}" +
  "*{box-sizing:border-box;}" +
  "body{margin:0;background:var(--paper);color:var(--canvas-ink);font-family:'Inter',sans-serif;" +
  "min-height:100vh;padding:0 0 60px;}" +
  ".topbar{background:var(--paper-raised);border-bottom:1px solid var(--line);padding:16px 28px;" +
  "display:flex;align-items:baseline;gap:10px;}" +
  ".brand-mark{font-family:'Fraunces',serif;font-size:19px;font-weight:600;color:var(--ink);letter-spacing:-0.02em;}" +
  ".brand-sub{font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--ink-faint);" +
  "letter-spacing:0.06em;text-transform:uppercase;}" +
  ".page-wrap{max-width:640px;margin:0 auto;padding:36px 24px;}" +
  ".page-title{font-family:'Fraunces',serif;font-size:28px;font-weight:500;color:var(--canvas-ink);" +
  "margin:0 0 24px;letter-spacing:-0.01em;}" +
  ".card{background:var(--paper-raised);border:1px solid var(--line);border-radius:14px;" +
  "padding:24px;box-shadow:0 1px 2px rgba(31,41,55,0.04);margin-bottom:16px;}" +
  ".meta{font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--ink-faint);" +
  "letter-spacing:0.06em;text-transform:uppercase;margin-bottom:6px;}" +
  ".card-title{font-family:'Fraunces',serif;font-size:20px;font-weight:500;color:var(--ink);margin:0 0 18px;}" +
  ".field{margin-bottom:16px;}" +
  ".field-label{display:block;font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--ink-faint);" +
  "letter-spacing:0.04em;text-transform:uppercase;margin-bottom:6px;}" +
  ".flag-new{display:inline-block;background:var(--amber-soft);color:var(--amber);" +
  "font-family:'IBM Plex Mono',monospace;font-size:10px;text-transform:uppercase;letter-spacing:0.04em;" +
  "border-radius:5px;padding:2px 6px;margin-left:6px;}" +
  ".input{width:100%;font-family:'Inter',sans-serif;font-size:14px;color:var(--ink);" +
  "background:#fff;border:1px solid var(--line);border-radius:8px;padding:9px 11px;}" +
  ".input:disabled{background:var(--paper-recessed);color:var(--ink-faint);}" +
  "textarea.input{resize:vertical;}" +
  ".rationale{font-style:italic;color:var(--ink-faint);font-size:13px;margin:0;}" +
  ".btn{font-family:'Inter',sans-serif;font-size:14px;font-weight:600;color:#fff;background:var(--amber);" +
  "border:none;border-radius:8px;padding:11px 20px;cursor:pointer;}" +
  ".btn:hover{opacity:0.92;}" +
  ".checkbox-row{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--ink-faint);}" +
  ".result{font-family:'IBM Plex Mono',monospace;font-size:12px;color:var(--amber);margin-top:12px;}" +
  ".empty{color:var(--canvas-ink-faint);}";

function htmlPage_(title, bodyHtml) {
  return (
    "<!DOCTYPE html><html><head><base target=\"_top\">" +
    "<meta charset=\"UTF-8\">" +
    "<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">" +
    "<link href=\"https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap\" rel=\"stylesheet\">" +
    "<style>" + REVIEW_PAGE_CSS_ + "</style></head><body>" +
    "<div class=\"topbar\"><span class=\"brand-mark\">Threadline</span>" +
    "<span class=\"brand-sub\">Action Item Review</span></div>" +
    "<div class=\"page-wrap\"><h1 class=\"page-title\">" + escapeHtml_(title) + "</h1>" + bodyHtml + "</div>" +
    "</body></html>"
  );
}

/**
 * Review page: shown when the email's per-item "Review" link is opened.
 * "Edit before accepting" unlocks every field, including thread/sub-thread
 * placement itself: pick an existing thread from a dropdown (with "+ Add New
 * Thread" as the last option, revealing a name box), then that thread's
 * existing sub-threads (with its own "+ Add New Subthread" option) - unless
 * a brand-new thread was picked, in which case there's nothing to list yet
 * so the sub-thread name box appears directly, no dropdown. Submitting
 * commits straight to the board and redirects to the Work Tracker - there's
 * no separate commit screen, since every field (including placement) is
 * already editable right here.
 */
function renderItemReviewPage_(token) {
  const row = getPendingActionRow_(token);
  if (!row) return htmlPage_("Not found", "<p class=\"empty\">No pending action item found for this link.</p>");
  if (row.Status !== "proposed") {
    return htmlPage_("Already reviewed", "<p class=\"empty\">This item was already marked <b>" + escapeHtml_(row.Status) + "</b>.</p>");
  }

  const board = getBoardData().board;
  const boardMap = {};
  board.forEach(function (t) {
    boardMap[t.name] = t.subthreads.map(function (s) {
      return { name: s.name, tag: s.tag };
    });
  });
  const threadNames = Object.keys(boardMap);
  const threadMatches = threadNames.indexOf(row.ThreadName) !== -1;
  const useNewThread = row.IsNewThread || !threadMatches;
  const useNewSubThread = useNewThread || row.IsNewSubThread;
  const webAppUrl = getWebAppBaseUrl_() || "";

  const threadOptions = threadNames
    .map(function (name) {
      const selected = !useNewThread && name === row.ThreadName ? " selected" : "";
      return "<option value=\"" + escapeHtml_(name) + "\"" + selected + ">" + escapeHtml_(name) + "</option>";
    })
    .join("");

  const body =
    "<div class=\"card\">" +
    "<div class=\"meta\">From meeting</div>" +
    "<h3 class=\"card-title\">" + escapeHtml_(row.MeetingTitle) + "</h3>" +

    "<div class=\"field\"><label class=\"field-label\">Decision</label>" +
    "<select id=\"decision\" class=\"input\" onchange=\"toggleEdit()\">" +
    "<option value=\"accept\">Accept as-is</option>" +
    "<option value=\"edit\">Edit before accepting</option>" +
    "</select></div>" +

    "<div class=\"field\"><label class=\"field-label\">Item text</label>" +
    "<textarea id=\"text\" class=\"input\" rows=\"3\" disabled>" + escapeHtml_(row.Text) + "</textarea></div>" +

    "<div class=\"field\"><label class=\"field-label\">Owner</label>" +
    "<input id=\"owner\" class=\"input\" type=\"text\" disabled value=\"" + escapeHtml_(row.Owner) + "\"></div>" +

    "<div class=\"field\"><label class=\"field-label\">Opening date</label>" +
    "<input id=\"openDate\" class=\"input\" type=\"date\" disabled value=\"" + escapeHtml_(row.OpenDate) + "\"></div>" +

    "<div class=\"field\"><label class=\"field-label\">Closing date</label>" +
    "<input id=\"closeDate\" class=\"input\" type=\"date\" disabled value=\"" + escapeHtml_(row.OpenDate) + "\"></div>" +

    "<div class=\"field\"><label class=\"field-label\">Thread" +
    (row.IsNewThread ? "<span class=\"flag-new\">proposed new</span>" : "") + "</label>" +
    "<select id=\"threadSelect\" class=\"input\" disabled onchange=\"onThreadChange()\">" +
    threadOptions +
    "<option value=\"__new__\"" + (useNewThread ? " selected" : "") + ">+ Add New Thread</option>" +
    "</select>" +
    "<input id=\"newThreadName\" class=\"input\" type=\"text\" placeholder=\"New thread name\" disabled " +
    "value=\"" + escapeHtml_(row.ThreadName) + "\" " +
    "style=\"display:" + (useNewThread ? "block" : "none") + ";margin-top:8px;\"></div>" +

    "<div class=\"field\" id=\"subThreadDropdownField\" style=\"display:" + (useNewThread ? "none" : "block") + "\">" +
    "<label class=\"field-label\">Sub-thread" +
    (row.IsNewSubThread ? "<span class=\"flag-new\">proposed new</span>" : "") + "</label>" +
    "<select id=\"subThreadSelect\" class=\"input\" disabled onchange=\"onSubThreadChange()\"></select></div>" +

    "<div class=\"field\" id=\"newSubThreadField\" style=\"display:" + (useNewSubThread ? "block" : "none") + "\">" +
    "<label class=\"field-label\">" + (useNewThread ? "Sub-thread name (thread is new, so no list to pick from)" : "New sub-thread name") + "</label>" +
    "<input id=\"newSubThreadName\" class=\"input\" type=\"text\" placeholder=\"Sub-thread name\" disabled " +
    "value=\"" + escapeHtml_(row.SubThreadName) + "\">" +
    "<input id=\"newSubThreadTag\" class=\"input\" type=\"text\" placeholder=\"Tag (optional)\" disabled " +
    "value=\"" + escapeHtml_(row.SubThreadTag) + "\" style=\"margin-top:8px;\"></div>" +

    "<div class=\"field\"><label class=\"field-label\">Why the model chose this</label>" +
    "<p class=\"rationale\">" + escapeHtml_(row.Rationale) + "</p></div>" +

    "<button class=\"btn\" onclick=\"submitDecision()\">Submit</button>" +
    "<p id=\"result\" class=\"result\"></p>" +
    "</div>" +

    "<script>" +
    "var boardMap=" + JSON.stringify(boardMap) + ";" +
    "var PROPOSED_SUB_NAME=" + JSON.stringify(row.SubThreadName) + ";" +
    "var PROPOSED_IS_NEW_SUB=" + (row.IsNewSubThread ? "true" : "false") + ";" +

    "function populateSubThreadOptions(selectedName,isNewSub){" +
    "var sel=document.getElementById('subThreadSelect');" +
    "var threadVal=document.getElementById('threadSelect').value;" +
    "var subs=boardMap[threadVal]||[];" +
    "sel.innerHTML='';" +
    "subs.forEach(function(s){" +
    "var opt=document.createElement('option');opt.value=s.name;opt.textContent=s.name;" +
    "if(!isNewSub&&s.name===selectedName)opt.selected=true;" +
    "sel.appendChild(opt);});" +
    "var newOpt=document.createElement('option');newOpt.value='__new__';newOpt.textContent='+ Add New Subthread';" +
    "var exists=subs.some(function(s){return s.name===selectedName;});" +
    "if(isNewSub||!exists)newOpt.selected=true;" +
    "sel.appendChild(newOpt);" +
    "}" +

    "function onThreadChange(){" +
    "var isNewThread=document.getElementById('threadSelect').value==='__new__';" +
    "document.getElementById('newThreadName').style.display=isNewThread?'block':'none';" +
    "document.getElementById('subThreadDropdownField').style.display=isNewThread?'none':'block';" +
    "if(isNewThread){" +
    "document.getElementById('newSubThreadField').style.display='block';" +
    "}else{" +
    "populateSubThreadOptions('',false);" +
    "onSubThreadChange();" +
    "}" +
    "}" +

    "function onSubThreadChange(){" +
    "var isNewSub=document.getElementById('subThreadSelect').value==='__new__';" +
    "document.getElementById('newSubThreadField').style.display=isNewSub?'block':'none';" +
    "}" +

    "function initForm(){" +
    "var isNewThread=document.getElementById('threadSelect').value==='__new__';" +
    "if(!isNewThread){" +
    "populateSubThreadOptions(PROPOSED_SUB_NAME,PROPOSED_IS_NEW_SUB);" +
    "onSubThreadChange();" +
    "}" +
    "}" +
    "initForm();" +

    "function toggleEdit(){" +
    "var isEdit=document.getElementById('decision').value==='edit';" +
    "['text','owner','openDate','closeDate','threadSelect','newThreadName','subThreadSelect','newSubThreadName','newSubThreadTag']" +
    ".forEach(function(id){document.getElementById(id).disabled=!isEdit;});" +
    "}" +

    "function submitDecision(){" +
    "document.getElementById('result').innerText='Committing...';" +
    "var isNewThread=document.getElementById('threadSelect').value==='__new__';" +
    "var isNewSubThread=isNewThread||document.getElementById('subThreadSelect').value==='__new__';" +
    "var threadName=isNewThread?document.getElementById('newThreadName').value:document.getElementById('threadSelect').value;" +
    "var subThreadName=isNewSubThread?document.getElementById('newSubThreadName').value:document.getElementById('subThreadSelect').value;" +
    "var subThreadTag=isNewSubThread?document.getElementById('newSubThreadTag').value:'';" +
    "google.script.run" +
    ".withSuccessHandler(function(){" +
    "document.getElementById('result').innerText='Committed to your Work Tracker. Redirecting...';" +
    (webAppUrl ? "top.location.href=" + JSON.stringify(webAppUrl) + ";" : "") +
    "})" +
    ".withFailureHandler(function(err){document.getElementById('result').innerText='Error: '+err.message;})" +
    ".submitItemDecision(" + JSON.stringify(token) + "," +
    "document.getElementById('text').value," +
    "document.getElementById('owner').value," +
    "document.getElementById('openDate').value," +
    "document.getElementById('closeDate').value," +
    "threadName,isNewThread,subThreadName,isNewSubThread,subThreadTag);" +
    "}" +
    "</script>";

  return htmlPage_("Review action item", body);
}

/**
 * Called from the Review page's Submit. Fields always reflect the final
 * intended values (unedited fields simply still hold the original
 * proposal), so this just finds-or-creates the chosen thread/sub-thread and
 * writes the item straight to the board via the same
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

/**
 * Runs the whole pipeline end to end: poll -> extract -> create pending rows
 * -> email. Intended as the single function the 15-minute trigger calls -
 * see installActionItemPipelineTrigger().
 */
function runActionItemPipeline() {
  pollMeetingNotes();
  extractActionItemsForFetchedMeetings();
  createPendingActionsForExtractedMeetings();
  sendPendingActionEmails();
}

/**
 * Installs (or reinstalls) the 15-minute trigger for the full pipeline.
 * Replaces installMeetingPollTrigger() now that all four pieces are wired
 * together - run this once instead, after setting NIM_API_KEY, NIM_MODEL,
 * and (once deployed) WEBAPP_URL in Script Properties.
 */
function installActionItemPipelineTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === "runActionItemPipeline" || fn === "pollMeetingNotes") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("runActionItemPipeline").timeBased().everyMinutes(15).create();
  Logger.log("Installed 15-minute runActionItemPipeline trigger.");
}

/**
 * One-off cleanup for the duplicate PendingActions rows caused by the
 * pre-lock race condition (overlapping runs both emailing/creating rows for
 * the same meeting). Groups rows by everything that would be identical for
 * a genuine duplicate (EventID/Text/Owner/ThreadName/SubThreadName), keeps
 * the earliest row in each group, and deletes the rest - but ONLY among
 * rows still at Status "proposed", since those were never acted on and
 * nothing on the board depends on them.
 *
 * Duplicates found at any OTHER status (e.g. "committed") are just logged,
 * not touched - if you clicked two different duplicate tokens for the same
 * item, there's a real duplicate item sitting on the Work Tracker board
 * too, and only you can tell which copy (if either) to delete there.
 * Safe to run multiple times; logs exactly what it deleted.
 */
function dedupePendingActions() {
  const sheet = getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  const groups = {};
  for (var r = 1; r < values.length; r++) {
    const row = values[r];
    const key = [
      row[col.EventID], row[col.Text], row[col.Owner],
      row[col.ThreadName], row[col.SubThreadName],
    ].join("||");
    if (!groups[key]) groups[key] = [];
    groups[key].push({ rowIndex: r + 1, status: row[col.Status], createdAt: row[col.CreatedAt] });
  }

  var rowsToDelete = [];
  var loggedOtherStatusDupes = 0;

  Object.keys(groups).forEach(function (key) {
    const group = groups[key];
    if (group.length < 2) return;

    const proposedRows = group.filter(function (g) { return g.status === "proposed"; });
    const otherRows = group.filter(function (g) { return g.status !== "proposed"; });

    if (proposedRows.length > 1) {
      proposedRows.sort(function (a, b) { return a.createdAt - b.createdAt; });
      // Keep the earliest, delete the rest.
      proposedRows.slice(1).forEach(function (g) { rowsToDelete.push(g.rowIndex); });
    }

    if (otherRows.length > 1) {
      loggedOtherStatusDupes++;
      Logger.log(
        'Found %s non-"proposed" duplicate row(s) for "%s" at rows %s - not auto-deleted. ' +
        "Check your Work Tracker board for a matching duplicate item and remove it there if present.",
        otherRows.length, key.split("||")[1], otherRows.map(function (g) { return g.rowIndex; }).join(", ")
      );
    }
  });

  // Delete highest row index first so earlier indices stay valid.
  rowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (rowIndex) {
    sheet.deleteRow(rowIndex);
  });

  Logger.log("Deleted %s duplicate 'proposed' row(s). %s other-status duplicate group(s) need a manual look (see above).",
    rowsToDelete.length, loggedOtherStatusDupes);
}
