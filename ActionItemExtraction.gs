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
 * item), each waiting there with Status "proposed" until reviewed.
 *
 * Piece 4/4: no email, no separate page - Index.html's "Action Item
 * Proposals" tab lists every "proposed" row (getPendingProposals()) with
 * everything editable inline - text, owner, opening/closing date, and the
 * thread/sub-thread it lands in. Its Commit button calls submitItemDecision()
 * to write straight to the board via the same addThread/addSubThread/addItem
 * functions Code.gs already uses; its Dismiss button calls dismissProposal()
 * to discard a proposal without touching the board.
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
      // Anything past "waiting" has already had its notes doc found (or
      // further processed) - skip it for good. The old check here only
      // recognized "notes_fetched" as done, so once a meeting advanced to
      // "extracted"/"rows_created" the next poll no longer saw it as
      // handled, found the same still-present attachment, and reset it
      // back to "notes_fetched" - re-triggering extraction and proposal
      // creation for meetings that were already fully processed. That was
      // the actual source of the repeated/duplicate proposals.
      if (existing && existing.Status !== "waiting") return;

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

// ---- Piece 3/4: PendingActions tab ----

/**
 * Expands each "extracted" meeting's ProposalsJson into one PendingActions
 * row per item, then advances the meeting to "rows_created" - its items are
 * now sitting there with Status "proposed", ready for the Action Item
 * Proposals tab to list.
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

// ---- Piece 4/4: in-app proposal review (Index.html's "Action Item
// Proposals" tab) ----
// No email, no separate standalone page - proposals sit in PendingActions
// until reviewed inline in the tracker itself. getPendingProposals() feeds
// the tab's list; submitItemDecision()/dismissProposal() are what its
// Commit/Dismiss buttons call.

function getPendingActionRow_(token) {
  const rows = rowsToObjects_(getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS));
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].Token === token) return rows[i];
  }
  return null;
}

/**
 * Returns every "proposed" item for the Action Item Proposals tab to render.
 */
function getPendingProposals() {
  return rowsToObjects_(getSheet_(ACTION_ITEM_SHEET_NAMES.PENDING_ACTIONS)).filter(function (r) {
    return r.Status === "proposed";
  });
}

/**
 * Called from the tab's Dismiss button - discards a proposal without
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
 * Called from the tab's Commit button. Fields always reflect the final
 * intended values, so this just finds-or-creates the chosen thread/
 * sub-thread and writes the item straight to the board via the same
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
 * Runs the whole pipeline end to end: poll -> extract -> create pending
 * rows. No email step - proposals just sit in PendingActions until reviewed
 * in the Action Item Proposals tab. Intended as the single function the
 * 15-minute trigger calls - see installActionItemPipelineTrigger().
 */
function runActionItemPipeline() {
  pollMeetingNotes();
  extractActionItemsForFetchedMeetings();
  createPendingActionsForExtractedMeetings();
}

/**
 * Installs (or reinstalls) the 15-minute trigger for the full pipeline.
 * Replaces installMeetingPollTrigger() now that all pieces are wired
 * together - run this once instead, after setting NIM_API_KEY and
 * NIM_MODEL in Script Properties.
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
