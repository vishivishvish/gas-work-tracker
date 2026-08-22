# Google Apps Script Work Tracker

A horizontal, Trello-inspired work tracker with a high degree of flexibility & extensibility for managing a large number of ongoing
work threads. Each numbered thread branches out into sub-threads
laid out side by side, and each sub-thread holds a list of lettered,
checkable action steps. Threads can be dragged to reorder them, colored
from a 31-option picker, and threads/sub-threads/items can each carry an
optional date. A second tab pivots the same data into a per-person
checklist of open items.

Everything is read/write straight from the underlying Sheet, so nothing
about the board's state ever lives only in the browser tab.

Runs as a Google Apps Script web app, backed by a Google Sheet, inside your
own Google Workspace account: no external hosting, no separate login
required beyond your own Google account.

## Features

Keep this list current whenever a feature is added or changed - it's the
one place that should always reflect what the app actually does today.

- **Threads, sub-threads, and action items** - a horizontal thread board
  where each thread expands into side-by-side sub-threads, each holding a
  checkable, lettered list of action steps.
- **Thread numbering** - threads are numbered by their on-page position
  (computed at render time, never stored), so numbers always stay correct
  after adding, deleting, or reordering.
- **Thread drag-and-drop reordering** - drag a thread's header up or down
  to reorder it relative to other threads; persisted via `reorderThreads()`.
- **Sub-thread drag-and-drop reordering** - drag a sub-thread's header
  left or right to reorder it within its own thread (cross-thread drops are
  blocked); persisted via `reorderSubThreads()`.
- **Optional dates** - threads and sub-threads can each carry an optional
  opened date, set via the same prompts used for naming. Items carry two
  dates: an opening date and a closing date - leaving closing date equal to
  opening date marks an item still open/WIP, and updating it to the real
  date it closed marks it done. This is independent of the checkbox/
  strikethrough, which is just a visual toggle.
- **Per-thread color picker** - each thread has a color-picker swatch
  offering 31 light, named colors ordered in a VIBGYOR gradient (violet
  through red, with neutrals at the end); hovering a swatch shows its name.
  Picking a color re-themes that thread's card background, text, and every
  nested sub-thread/item together via CSS custom properties scoped to that
  thread - a sub-thread's tag always renders in a color family distinct
  from its body text for visibility, regardless of which color is chosen.
- **"Threads for Each Person" tab** - a second tab that pivots the same
  board data into one checklist per owner, listing every *open* (unchecked)
  item assigned to them across all threads/sub-threads, with the source
  thread name and tag shown for context. Items can be checked off directly
  from this view. Computed entirely client-side from the already-loaded
  board data - no extra backend call. Each person's card is collapsible
  (same chevron pattern as threads), though the open/closed state resets on
  reload since there's no "owner" row in the Sheet to persist it against.
- **Live polling** - the UI polls `getLastModified()` every 5 seconds and
  only re-fetches the full board when it's actually changed, so multiple
  open tabs (or direct Sheet edits, with the installable trigger) stay in
  sync without a manual refresh.
- **Direct-Sheet-edit sync** - with the optional installable `onEdit`
  trigger, edits made straight in the Google Sheet (not through the web
  app) also bump `LastModified` and show up on the next poll.

## Roadmap: automated action-item extraction from meeting notes

In progress - built in four pieces, in `ActionItemExtraction.gs`, tracked via
a new `ProcessedMeetings` tab on the same Sheet.

**Piece 1 (done): poll for the Gemini notes doc.** Every meeting on the
calendar gets Gemini-generated notes attached to its Calendar event once the
meeting ends (timing varies with meeting length). `pollMeetingNotes()` runs
on a 15-minute time-driven trigger, sweeps events from the last
`MEETING_LOOKBACK_HOURS` (24) hours - a sliding window recomputed from "now"
on every run, not tracked against the previous run - via the Calendar
Advanced Service, and marks each one `waiting` or `notes_fetched` in
`ProcessedMeetings` depending on whether that attachment has shown up yet, so
a meeting is only ever read once, whenever its notes actually land. Once a
meeting falls outside that 24-hour window it's no longer checked at all, so a
meeting whose notes take longer than that to appear (or one from further
back than 24 hours ago that was never caught) won't be picked up
automatically - bump `MEETING_LOOKBACK_HOURS` and run `pollMeetingNotes`
manually once to backfill a case like that. The notes doc's own creation
date (its file metadata, not the calendar event) is captured here too, as
`NotesDate` - this becomes every action item's default opening date.

**Piece 2 (done): propose action items via NIM.** For every `notes_fetched`
meeting, `extractActionItemsForFetchedMeetings()` sends its notes text, plus
a summary of the board's current threads/sub-threads, to NVIDIA NIM's
OpenAI-compatible chat completions endpoint (model/API key set as Script
Properties, never hardcoded). The model returns a JSON array of proposed
items - text, owner, target thread/sub-thread (existing or new), and a short
rationale for that placement - stored as-is on the meeting's
`ProcessedMeetings` row (`ProposalsJson` column, `Status: extracted`).
Nothing on the board changes yet.

**Piece 3 (done): pending-approval tab.**
`createPendingActionsForExtractedMeetings()` expands each `extracted`
meeting's `ProposalsJson` into one row per item on a new `PendingActions` tab
(each row gets its own token and the meeting's `NotesDate` as its default
`OpenDate`, `Status: proposed`), advancing the meeting to `rows_created`.
No email - proposals just sit there until reviewed in the tab described
below.

**Piece 4 (done): in-app review, no email.** A third tab in `Index.html`,
"Action Item Proposals" (with a count badge, refreshed every 30s), lists
every `proposed` row via `getPendingProposals()`, sharing the board's visual
language since it's literally the same page. Every field is directly
editable inline - text, owner, opening/closing date (`<input type="date">`,
both defaulting to the meeting's notes date, matching the open==close means
still-open/WIP convention) - and thread/sub-thread placement itself: a
Thread dropdown (existing threads, plus "+ Add New Thread" which reveals a
name box), then that thread's existing sub-threads in a second dropdown
(plus its own "+ Add New Subthread", revealing a name + tag box) - unless a
brand-new thread was picked, in which case there's no existing list to show,
so the sub-thread name box appears directly.

**Commit** calls `findOrCreateThread_`/`findOrCreateSubThread_` (matching
existing names case-insensitively before creating anything new) and the
existing `addItem()` to write straight to the board via `submitItemDecision()`,
marking the row `committed` and reloading both the proposals list and the
board. **Dismiss** calls `dismissProposal()` to mark a row `rejected` without
touching the board.

`runActionItemPipeline()` chains poll → extract → create pending rows for the
15-minute trigger (`installActionItemPipelineTrigger()` installs it,
replacing the piece-1-only `installMeetingPollTrigger()`). Setup: set
`NIM_API_KEY` and `NIM_MODEL` in Script Properties, deploy the web app, then
run `installActionItemPipelineTrigger()` once.

### Duplicate-proposal bug and fixes

Early versions of this pipeline sent an email per proposal and re-created
duplicate rows/emails for the same meeting. Three separate things were
wrong, now all fixed:

1. **The actual root cause**: `pollMeetingNotes()`'s skip-check only
   recognized `Status === "notes_fetched"` as "already handled." Once a
   meeting advanced further (`extracted`, `rows_created`), the *next* poll no
   longer saw it as done, found the same still-present Gemini attachment, and
   reset its status back to `notes_fetched` - re-triggering extraction and
   proposal creation for a meeting that was already fully processed, every 15
   minutes, for as long as it stayed within the lookback window. Fixed: the
   check now skips anything past `waiting`, full stop.
2. **A real but secondary race**: two overlapping runs (the trigger firing
   while a manual test run was still in flight, say) could both read a
   meeting as not-yet-handled before either wrote its new status back. Fixed:
   each of the three stage functions holds `LockService`'s script lock for
   its whole run (`withScriptLock_()`) - an overlapping run just skips.
3. **Defense in depth**: `createPendingActionsForExtractedMeetings()` also
   checks directly whether `PendingActions` already has any row for a
   meeting's `EventID` (the Calendar event's own unique ID) before creating
   more, regardless of what `ProcessedMeetings.Status` claims - idempotent
   per-meeting even if status ever drifts out of sync some other way.

`dedupePendingActions()` is the one-off cleanup for duplicate rows that
existed before these fixes landed.

## Files

- `Code.gs` - backend: Sheet schema setup, `doGet()` web app entry point,
  and all read/write functions (add/rename/delete/toggle for threads,
  sub-threads, and items).
- `Index.html` - frontend: the board UI, rendered from `getBoardData()`,
  polling every 5 seconds for changes and re-rendering when the Sheet's
  `LastModified` timestamp moves.
- `thread-manager-mockup-v2.html` - the original static design mockup
  (kept for reference; not used by the running app).

## Data model (Google Sheet, auto-created by `setupSheets()`)

- `Threads`: ThreadID, Name, Order, Collapsed, DateOpened, Color
- `SubThreads`: SubThreadID, ThreadID, Name, Tag, Order, Collapsed, DateOpened
- `Items`: ItemID, SubThreadID, Text, Checked, Owner, Order, OpenDate, CloseDate
- `Meta`: Key, Value (holds `LastModified`, bumped by every write)

Letters (A, B, C...) for action steps and numbers (1, 2, 3...) for threads
are both computed at read time from each row's position - never stored -
so deleting/adding/reordering never causes duplicate or stale labels.

The Sheet is referenced by its permanent Drive file ID (stored in this
script's Script Properties), never by folder path, so it can be moved
between Drive folders at any time without breaking anything.

## One-time setup

1. Go to [script.google.com/home](https://script.google.com/home) > New project.
2. Paste `Code.gs`'s contents into the default `Code.gs` file.
3. File > New > HTML file, name it exactly `Index`, paste `Index.html`'s contents in.
4. Select `setupSheets` in the function dropdown (top toolbar) and click **Run**.
   First run creates a new Sheet called "Threadline Data", builds all four
   tabs with headers, and stores its ID in this script's Script Properties.
   Approve the permissions prompt. Check `Execution Log` (top toolbar) for the Sheet's URL.
5. *If you had this project set up before the `DateOpened`/`Date` columns
   existed*, select `migrateAddDateColumns` in the function dropdown and
   click **Run** once. It only adds the missing columns and never touches
   existing data - brand-new setups already have them from `setupSheets`
   and can skip this.
6. *If you had this project set up before per-thread `Color` existed*,
   select `migrateAddThreadColor` in the function dropdown and click **Run**
   once, for the same reason as above.
6b. *If you had this project set up before items had separate opening/closing
   dates (a single `Date` column instead)*, select
   `migrateItemDatesToOpenClose` in the function dropdown and click **Run**
   once - it renames `Date` to `OpenDate` in place (keeping existing values)
   and backfills `CloseDate` to match `OpenDate` for every existing item, so
   they all start out looking "still open," same as a freshly-added item
   would. Brand-new setups already get both columns from `setupSheets`.
7. *(Optional, recommended)* Set up direct-edit sync: select `installEditTrigger`
   in the function dropdown and click **Run**. (Not done via the Triggers UI's
   "Add Trigger" dialog - its "From spreadsheet" event source is only offered
   to scripts *bound* to a Sheet, i.e. opened via Extensions > Apps Script
   from inside the Sheet itself. Since this is a standalone script, that
   option won't appear there, so the trigger is installed in code instead.)
   Without this, edits made directly in the Sheet (rather than through the
   web app) won't be picked up until the next write through the app bumps
   `LastModified` - with it, direct Sheet edits sync too.
8. **Deploy > New deployment > Web app**. Execute as: **Me**. Who has
   access: **Only myself** (or your Workspace domain, if teammates should
   use it too - they'd then also need edit access to the underlying Sheet).
   Copy the resulting web app URL - that's the bookmark you'll use daily.
9. To ship code changes later **without changing that URL**: **Deploy >
   Manage deployments** > pick the existing deployment > pencil/Edit icon >
   Version: **New version** > Deploy. Creating a brand-new deployment
   instead of editing the existing one gives you a different URL each time.

## How live updates work

Since Apps Script web apps can't push updates over a server socket,
`Index.html` instead polls `getLastModified()` every 5 seconds (a single
cheap cell read), only re-fetching the full board (`getBoardData()`) once
that timestamp has actually moved. Every write function bumps the timestamp, and (with the
installable trigger from step 7) so does any direct edit in the Sheet.

## Thread and sub-thread reordering

Threads are numbered by their on-page position, and each thread's header
(the row with the chevron, number, and title) is a drag handle: drag one
thread and drop it above or below another to reorder. The drop calls
`reorderThreads()`, which rewrites the `Order` column for every thread, and
the board re-renders immediately with updated numbers - other open tabs
pick up the new order on their next 5-second poll.

Sub-threads work the same way but horizontally: each sub-thread's header is
also a drag handle, and dropping it before or after another sub-thread
*within the same thread* calls `reorderSubThreads()` to persist the new
order. Dragging a sub-thread onto a different thread is a no-op - reorder
is scoped to one thread at a time.

## Dates

Threads and sub-threads each carry an optional opened date, set through the
same prompts used for naming (leave blank to clear it).

Items carry two dates - opening and closing - both asked for (as sequential
prompts, closing pre-filled with whatever was just typed for opening) when
adding or editing an item. Leaving closing date equal to opening date means
the item is still open/WIP; once it's actually done, edit the item and change
the closing date to the real date it closed. This is separate from the
checkbox/strikethrough, which is just a visual "done" toggle - the two dates
are the source of truth for when something opened and closed.

## Thread colors

Each thread has a small color swatch in its header. Clicking it opens a
picker with 31 light, named colors laid out in VIBGYOR order (violet
through red, with gray/graphite/cocoa/birch as neutrals at the end);
hovering any swatch shows its name as a tooltip. Selecting a color calls
`setThreadColor()`, which persists it to the `Color` column, and the whole
thread - card background, title, dates, progress label, and every nested
sub-thread and item - re-themes together via CSS custom properties scoped
to that thread's element. A sub-thread's tag is always rendered in a
different color family (teal vs. amber) than the rest of its text, so it
stays visually distinct no matter which of the 31 thread colors is active.

## Threads for Each Person

A second tab, next to the main "Work Tracker" tab, pivots the same board
data by owner: for every person named in an item's "owner" field, it lists
a checklist of all of that person's *open* (unchecked) items across every
thread and sub-thread, with the originating thread's name and tag shown
alongside each item for context. Checking an item off here calls the same
`toggleItem()` used on the main board, so it disappears from this view (and
gets marked done everywhere else) immediately. This view is computed
entirely in the browser from the board data already loaded for the main
tab - no extra call to the backend.

Each person's card has a chevron and collapses/expands the same way a
thread does. Unlike thread/sub-thread collapse state, this isn't persisted
to the Sheet - there's no "owner" row to store it against, since owners are
derived from item data rather than being their own entity - so it resets to
all-open on the next page load.
