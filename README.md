# Google Apps Script Work Tracker

A horizontal, Trello-like work tracker with a high degree of flexibility & extensibility for managing a large number of ongoing
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

## Meeting action-item extraction (manual, in `ActionItemExtraction.gs`)

An earlier version of this tried to run fully automatically (a 15-minute
polling trigger that emailed you proposals). It kept producing duplicates -
a bug in the polling logic re-triggered extraction for meetings that had
already been processed - so the design changed to fully manual/button-driven
instead: nothing runs on a schedule, and a given meeting can only ever be
extracted once, ever, by construction (see below), not by relying on getting
some background process's bookkeeping right.

**"Meetings" tab.** `getCalendarMeetingsWithAttachments()` lists every
Calendar event since `MEETINGS_SINCE_DATE` (currently 2026-05-19) with at
least one attachment, reverse-chronological, never later than right now
regardless of what's still ahead on the calendar. Each meeting shows its
date/time, attendees, attachment names, and is left-bordered in that event's
own Calendar color (`Calendar.Colors.get().event`, falling back to the
primary calendar's default color if the event has none). An attachment is
flagged as a likely Gemini transcript if its title starts with "Notes" and
contains "Gemini" (`looksLikeGeminiTranscript_()`); a meeting only gets an
**Extract Action Items** button if at least one of its attachments matches.
This tab is deliberately **not** auto-polled - Calendar API calls are
heavier and quota-limited compared to the board's cheap Sheet-timestamp
check, and calendar history doesn't change every 5 seconds anyway - so it
loads once when opened, plus a manual Refresh button.

**Extracting.** Clicking **Extract Action Items** calls
`extractActionItemsForMeeting(eventId)`, which picks the best-matching
attachment (preferring the Notes+Gemini naming convention, falling back to
any Google Doc attachment), reads its text, sends it plus a summary of the
board's current threads/sub-threads to NVIDIA NIM's OpenAI-compatible chat
completions endpoint (model/API key as Script Properties, never hardcoded),
and writes one `PendingActions` row per proposed item (`Status: proposed`,
`OpenDate` defaulted from the notes doc's own creation date). It then
**permanently** records the `EventID` in a new `ExtractedMeetings` tab - the
Meetings tab checks that to swap the button to a grayed, disabled "Action
Items Extracted," and the server rejects a second extraction for the same
`EventID` even if the button were somehow clicked anyway. That's the whole
duplication guard now: no trigger, no polling, no window for a race.

**"Action Items Pending for Approval" tab.** Only present in the tab bar at
all while `getPendingProposals()` (any `PendingActions` row still `proposed`)
returns something - checked on the same 5-second poll the board already
uses, so the tab appears right after an extraction and disappears again once
every item from it has been accepted or rejected. Every field is directly
editable inline - text, owner, opening/closing date (`<input type="date">`,
both defaulting to the meeting's notes date, matching the open==close means
still-open/WIP convention) - and thread/sub-thread placement itself: a
Thread dropdown (existing threads, plus "+ Add New Thread" which reveals a
name box), then that thread's existing sub-threads in a second dropdown
(plus its own "+ Add New Subthread", revealing a name + tag box) - unless a
brand-new thread was picked, in which case there's no existing list to show,
so the sub-thread name box appears directly.

**Accept with Edits** calls `findOrCreateThread_`/`findOrCreateSubThread_`
(matching existing names case-insensitively before creating anything new)
and the existing `addItem()` to write straight to the board via
`submitItemDecision()`, marking the row `committed`. **Reject** calls
`dismissProposal()` to mark a row `rejected` without touching the board.

**Setup**: set `NIM_API_KEY` and `NIM_MODEL` in Script Properties, deploy
the web app. No trigger to install - there's nothing scheduled.

**`wipeActionItemPipeline()`** is a one-off reset: deletes the old
`ProcessedMeetings` tab entirely (belonged to the retired automatic-polling
design), clears `PendingActions` and `ExtractedMeetings` down to just their
headers, and removes any leftover trigger from that old design. It never
touches the Work Tracker board itself - anything already committed stays
exactly as is.

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
