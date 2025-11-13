# Kanban Alert – Google Sheets × LINE OA
<img width="1180" height="634" alt="image" src="https://github.com/user-attachments/assets/2b550fc2-7880-4047-acba-1fa51849ebb4" />

---

Google Apps Script project that turns a Google Sheet into a **mini Kanban notification system** with LINE integration.

It:

- Tracks **per-cell edits** on your Kanban sheet
- Keeps **daily snapshots** of all rows
- Sends:
  - **Nightly change summary** ("Daily Changes") to a LINE Official Account group
  - **Morning reminder** for tasks with *Next Due Date = today*
- Provides:
  - A **manual “Urgent” sidebar** to send selected changes
  - A **Change Log** dialog with filters, sorting, and CSV export
  - A **Trigger Status / Edit-window** dashboard

> 🕒 Timezone is fixed to **Asia/Bangkok** by default.

---

## Table of Contents

1. [Features](#features)  
2. [Architecture Overview](#architecture-overview)  
3. [Requirements](#requirements)  
4. [Installation](#installation)  
5. [Configuration](#configuration)  
6. [LINE OA Setup & Secrets](#line-oa-setup--secrets)  
7. [How It Works](#how-it-works)  
8. [UI & Usage](#ui--usage)  
9. [Triggers & Scheduling](#triggers--scheduling)  
10. [Data Model](#data-model)  
11. [Troubleshooting & FAQ](#troubleshooting--faq)  
12. [Contributing](#contributing)  
13. [License](#license)

---

## Features

### 1. Change Tracking

- Watches specific sheet tabs (e.g. `Nov.`) defined in `CONFIG.SHEETS`.
- Each row gets a **UID** stored in the **Task column’s Note**, so rows can move and still be matched.
- Every edit on watched columns is:
  - Stored in a per-day **changelog** (with timestamp & editor email)
  - Available in the **Change Log dialog** and for **nightly/morning summaries**

### 2. Daily Snapshots

- Keeps daily snapshots of all rows (serialized state) for each watched sheet.
- Uses `CONFIG.SNAPSHOT_KEEP` (default `45`) to keep only the most recent N days.
- Snapshots are used to compare **“start-of-day vs now”** to build meaningful change summaries.

### 3. LINE Notifications

- **Nightly summary** – runs `runNightlyAlert` once per day:
  - Only includes rows **“Acted today”** (`Acted Date`) and where `By` matches `Request by` or `Assigned to`.
  - Shows detailed per-field changes using diffs/snapshots.
- **Morning reminder** – runs `runDueTodaySend` once per day:
  - Lists tasks whose **Next Due Date == today** and `Progress != "Cancelled"`.
  - Shows badges like `overdue`, `due today`, `Xd left`.
- **Manual urgent send** – via sidebar:
  - You manually select rows & optionally add custom messages.
  - Sends a compact “Urgent: Today’s Changes” message.

### 4. Change Log UI

- Full-screen modal **Change Log** with:
  - Date range filters (Today / 7 days / 30 days / All)
  - Per-field, row, task, editor filters
  - Sorting (time, row, field, task, by, editor)
  - **Stats**: top fields edited & top editors
  - CSV export:
    - **Visible** (current filters)
    - **Full-range** (raw from Apps Script)

### 5. Edit Window & Lock

- You define a **Morning time** and **Night time**.
- When **Edit Lock** is ON:
  - Edits are **allowed only between Morning and Night**.
  - Outside the window:
    - Edits are **reverted** using a per-cell cache
    - User sees a **countdown modal** until unlock time

### 6. Developer Tools

- Debug helpers:
  - `debugInitializeRepairOnce()` – initialize system & triggers
  - `debugHardResetDailyTriggers()` – reset only the time-based triggers
  - `debugRepairAllTriggers()` – repair both daily & realtime triggers
  - `debugListTriggers()` – log current triggers & next run times
- Trigger Status UI:
  - `showTriggerStatus()` launches a **dashboard** to see & edit:
    - Morning & Night times
    - On/off status of nightly / morning / realtime triggers
    - Edit-lock toggle
    - “Run now” / “Send test” buttons

---

## Architecture Overview

All logic is in Google Apps Script (V8 runtime).

Key parts:

- **CONFIG** – global settings (headers, watched sheets, timezone, messaging).
- **Triggers**:
  - `runNightlyAlert` – nightly snapshot + summary send.
  - `runDueTodaySend` – morning “Next Due Date = today” reminder.
  - `onEdit` – realtime stamping + logging.
- **Snapshots & Changelog**:
  - `buildSnapshot_`, `saveSnapshotForDay_`, `loadLatestSnapshot_`
  - `appendChangelog_`, `getChangeLogData`, `getChangeLogCsv`
- **UI / HTML**:
  - `buildChangeLogDialogHtml_` – change log modal.
  - `buildLockedEditHtml_` – locked-edit countdown modal.
  - `buildSidebarHtml_` – manual urgent send sidebar.
  - `showTriggerStatus()` – trigger status dashboard.
- **LINE sending**:
  - `sendLineMulti_`, `sendLine_`, `sendLineOA_`, `sendLineNotify_`
- **Utilities**:
  - Date/time normalization (`normalizeDate_`, `normalizeTime_`, `displayDate_`, etc.)
  - Name matching (`splitNames_`, `byMatchesRow_`, `chooseBy_`)
  - Per-cell cache for reverting (`loadCellCache_`, `saveCellCache_`)

---

## Requirements

- A **Google account** with:
  - Access to [Google Sheets](https://docs.google.com/spreadsheets/)
  - Access to [Google Apps Script](https://script.google.com/)
- A **LINE Official Account** with:
  - Channel access token
  - Group ID for your notification group
- A Kanban-like Sheet with headers matching your `CONFIG.HEADERS`.

Default timezone: `Asia/Bangkok`  
Default header row: `2`

---

## Installation

### 1. Create/Prepare Your Sheet

1. Create a new Google Sheet (or use an existing one).
2. Create a tab (e.g. `Nov.`) and add the following headers on **row 2**:

   | Column       | Example header text |
   |-------------|---------------------|
   | A           | `Task`              |
   | B           | `Request by`        |
   | C           | `Assigned to`       |
   | D           | `Resources`         |
   | E           | `Start Date`        |
   | F           | `Acted Date`        |
   | G           | `Next Due Date`     |
   | H           | `Due Date`          |
   | I           | `Progress`          |
   | J           | `Note`              |
   | K           | `Meeting Time`      |
   | L           | `By`                |

> The **header text must match** what you configure in `CONFIG.HEADERS`.

### 2. Create Apps Script Project

1. In the Sheet: **Extensions → Apps Script**.
2. Delete any default code.
3. Create `main.gs` and paste in this repository's script (or `src/main.gs` if you split files).

### 3. Configure `CONFIG`

At the top of the script, you’ll see:

```js
var CONFIG = {
  SHEETS: ['Nov.'],
  HEADER_ROW: 2,
  SNAPSHOT_KEEP: 45,
  TITLE: 'Daily Changes',
  TIMEZONE: 'Asia/Bangkok',
  // ...
};

Adjust:

* `SHEETS`: list of tab names to watch.
* `HEADER_ROW`: row number where your headers are.
* `TIMEZONE`: usually `Asia/Bangkok` for this project.
* `TITLE`: text used in test messages.

### 4. Set Script Properties (Secrets)

**Important:** Do not hard-code real tokens in `CONFIG`.
Instead, this project expects them via **Script Properties**.

In Apps Script editor:

1. Go to **Project Settings → Script Properties**.

2. Add keys:

   * `OA_CHANNEL_ACCESS_TOKEN` – your LINE OA channel access token
   * `OA_GROUP_ID` – your LINE group ID (starts with `C...`)

3. Save.

(If you're also using LINE Notify, add `LINE_NOTIFY_TOKEN` similarly.)

### 5. Authorize & Initialize

1. In the Apps Script editor, select `debugInitializeRepairOnce` in the function dropdown.
2. Click ▶ Run.

   * The first run will prompt you to authorize the script.
3. After it finishes:

   * Row UIDs are stamped (as Notes in the Task column).
   * Baseline snapshots are created.
   * Triggers are ensured to be correct.

---

## Configuration

### CONFIG basics

```js
var CONFIG = {
  SHEETS: ['Nov.'],         // tabs to watch
  HEADER_ROW: 2,            // row containing column headers
  SNAPSHOT_KEEP: 45,        // days of snapshots to retain
  TITLE: 'Daily Changes',
  TIMEZONE: 'Asia/Bangkok',
  oauthScopes: [ ... ],
  DELIVERY_METHOD: 'OA_BROADCAST',
  PREVIEW_MODE: 'BY_CHANGE_DATE',
  HEADERS: { ... },
  WATCH_FIELDS: { ... },
  REQUIRE_BY_MATCH: true,
  SEND_IF_EMPTY: false,
  MAX_ROWS_PER_MESSAGE: 10,
  SEPARATOR_LINE: '------------------------------------',
  REALTIME_MIN_GAP_MS: 3000
};
```

#### `SHEETS`

* Which tabs are considered part of the Kanban.
* Only these sheets are scanned for snapshots, changelog, and notifications.

#### `HEADERS`

Mapping between logical field keys and your actual header texts:

```js
HEADERS: {
  task:        'Task',
  requestBy:   'Request by',
  assignedTo:  'Assigned to',
  resources:   'Resources',
  startDate:   'Start Date',
  changeDate:  'Acted Date',
  nextDueDate: 'Next Due Date',
  dueDate:     'Due Date',
  progress:    'Progress',
  note:        'Note',
  meetingTime: 'Meeting Time',
  by:          'By'
}
```

If your sheet uses different header names, change them here.

#### `WATCH_FIELDS`

Which fields trigger `Acted Date` & `By` stamping when edited:

```js
WATCH_FIELDS: {
  task:        false,
  resources:   false,
  startDate:   false,
  changeDate:  false,
  nextDueDate: false,
  note:        false,
  dueDate:     false,
  meetingTime: false,
  progress:    false,
  by:          false
}
```

Set `true` for any fields you want to treat as “acting on this task”.
Those edits will:

* Set `Acted Date` to now.
* Update `By` based on `Request by` / `Assigned to`.

#### Edit Lock & Daily sends

```js
CONFIG.ENABLE_DAILY = true;    // global switch for daily sends
```

* `dailyEnabled_()` / `setDailyEnabled(on)` wrap this with persisted settings.
* Edit window is determined by `getMorningTime_()` and `getNightTime_()`.

---

## LINE OA Setup & Secrets

### 1. Get Channel Access Token

1. Go to LINE Developers console.
2. Select your OA channel.
3. Issue a **Channel access token**.
4. Copy the token.

### 2. Get Group ID

Common approach:

* Temporarily log sender IDs via a debug endpoint in your bot, or use a tool / existing webhook to capture the `groupId` when the group sends a message.

Once you have something like `Cxxxxxxxxxxxxxxx`, use that as `OA_GROUP_ID`.

### 3. Save in Script Properties

In Apps Script → Project Settings → Script Properties:

| Key                       | Value (example)                     |
| ------------------------- | ----------------------------------- |
| `OA_CHANNEL_ACCESS_TOKEN` | `abc123...`                         |
| `OA_GROUP_ID`             | `Cdfe4608d5975c6fbb87d1e4b46ff21e9` |
| `LINE_NOTIFY_TOKEN`       | *(optional, if using Notify)*       |

### 4. Security Notes

* Treat all tokens as **secrets**.
* Never commit real tokens or group IDs to GitHub.
* If a token leaks:

  * **Revoke / regenerate** it in LINE Developers.
  * Update Script Properties with the new token.

---

## How It Works

### 1. onEdit Flow

When a user edits a watched sheet:

1. **Edit lock check**

   * If `EDIT_LOCK_ON` is true and current time is **outside** the edit window:

     * The edit is reverted using the per-cell cache.
     * A toast and (if possible) a modal countdown are shown.

2. **Changelog logging**

   * For each changed cell on watched fields:

     * Lookup field by header name.
     * Compare old vs new from per-cell cache.
     * Append an entry to `KANBAN_CHANGELOG_<sheetId>` with:

       * `ts`, `row`, `uid`, `field`, `from`, `to`, `task`, `by`, `who`.

3. **Acted Date / By stamping**

   * If the field is in `CONFIG.WATCH_FIELDS`:

     * Update `Acted Date` to now.
     * Compute `By`:

       * Prefer current `By` if it matches `Request by` or `Assigned to`.
       * Else use:

         * Single `Assigned to` if only one.
         * Single `Request by` if only one.
         * Else keep existing `By`.

4. **Cache update**

   * Update per-cell cache for future comparisons and revert logic.

### 2. Nightly Alert

`runNightlyAlert` (time-based trigger):

1. Guard:

   * Only run once per “custom day” (aligned with Morning time).
2. For each watched sheet:

   * Load baseline snapshot (yesterday at day-start).
   * Load current snapshot (now).
   * Load today’s changelog.
   * Determine rows that:

     * Were **acted today** (matching `Acted Date`), and
     * Have meaningful diffs (`diffLinesForRow_`) or per-field changes.
     * Optionally require `By` to match `Request by` / `Assigned to`.
   * Build a message section per row with:

     * Row number, task, progress
     * Request/Assigned/Resources/Start/Next Due/Deadline/Meeting/Note
     * Detailed change bullets.
   * Save today’s snapshot as the new baseline.
3. If there are rows or `SEND_IF_EMPTY = true`:

   * Send multi-part LINE message via `sendLineMulti_`.
4. Record that nightly send has completed for this day.

### 3. Morning Due Today Send

`runDueTodaySend`:

1. Guard:

   * Only run once per calendar day.
2. For each watched sheet:

   * Filter rows where:

     * `normalizeDate_(Next Due Date) == today`
     * `progress.toLowerCase() !== 'cancelled'`.
3. Build a “Good Morning” message with concise details per row.
4. Send via LINE OA if there are rows (or `SEND_IF_EMPTY = true`).

---

## UI & Usage

### 1. Google Sheets Menu

On spreadsheet open (`onOpen`), the script adds a custom menu:

**Kanban Alert**

* `▶ Manually Notification. 🔔`
  → opens the manual urgent-send sidebar.

* `▶ Preview Change Summary Notification. ☀️`
  → previews the **nightly** change summary (what would be sent tonight).

* `▶ Preview Tomorrow Task Notification. 🌙`
  → previews **tomorrow’s** “Next Due Date” list.

Additional debug / dev UIs (not always in menu, but callable):

* `openChangeLogDialog()`
  → opens the Change Log modal.

* `showTriggerStatus()`
  → opens the Trigger Status dashboard.

### 2. Manual “Urgent” Sidebar

* Open via menu: **Kanban Alert → Manually Notification. 🔔**
* Shows today’s acted rows with:

  * Sheet, row, task, status
  * Request / Start / Next Due / Deadline / badges
  * Change bullets
* You can:

  * Select / unselect rows.
  * Optionally write a custom text to replace the “Change” section for that row.
  * Click **Send selected** to push one combined LINE message.

### 3. Change Log Dialog

* Call `openChangeLogDialog()` (you can wire a menu item if desired).
* Features:

  * Quick range: Today / 7d / 30d / All
  * Manual `From` / `To` date
  * Filter by:

    * Field
    * Row #
    * Task text (contains)
    * By
    * Editor email (contains)
  * Auto-refresh: Off / 30s / 1m / 5m
  * CSV export:

    * Visible (current filters & sort)
    * Full range (raw for that date window)
  * Per-page pagination: 50 / 200 / 1000 / All
  * Highlighted difference (before/after) using `<mark>` for changed sections.

---

## Triggers & Scheduling

### Types of Triggers

* **Time-based**

  * `runDueTodaySend` – morning “Next Due Date” reminders.
  * `runNightlyAlert` – nightly snapshot & change summary.
* **Installable from Spreadsheet**

  * `onEdit` – realtime stamping & changelog.

### Managing Triggers

* **Automatic management**

  `ensureTriggersAlive_()`:

  * Deletes duplicates.
  * Respects `CONFIG.ENABLE_DAILY` and `dailyEnabled_()` toggle.
  * Ensures you have:

    * Exactly one `runNightlyAlert`.
    * Exactly one `runDueTodaySend`.
    * One realtime `onEdit`.

* **Manual debug helpers**

  * `debugHardResetDailyTriggers()`
    → Delete & recreate daily pair (morning/night).

  * `debugRepairAllTriggers()`
    → Clean morning/night + realtime `onEdit`.

  * `debugListTriggers()`
    → Log all triggers and next run times.

  * `hardResetDailyTriggers_()`
    → Internal helper to reset daily triggers.

### Edit Window

* Morning & Night times are stored as Script Properties:

  * `MORNING_TIME` (default `06:00`)
  * `NIGHT_TIME` (default `21:00`)

* Edit window logic:

  ```text
  allowed when Morning <= now <= Night  (same-day)
  ```

  If Night < Morning (rare), window “wraps” across midnight.

* **Custom day boundary** for snapshots & “today”:

  * Bound to Morning hour (e.g. 06:00).
  * Means “today” for acted rows lines up with your workday.

---

## Data Model

### Row Identification

* Each row gets a **UID** stored in the `Task` column’s **Note**.
* Functions:

  * `ensureRowUidsForSheet_()` – adds UIDs to all data rows.
  * This UID is used as the primary key for:

    * Snapshots
    * Changelog
    * Diff comparison

### Snapshots

Stored as Document Properties:

* Keys like:

  * `KANBAN_SNAPSHOT_<yyyy-MM-dd>_S<sheetId>`
  * `KANBAN_SNAPSHOT_LATEST_DATE_S<sheetId>`
  * `KANBAN_SNAPSHOT_PREV_DATE_S<sheetId>`
* Each snapshot is a JSON of:

  ```js
  {
    "<uid>": {
      uid, task, requestBy, assignedTo, resources,
      startDate, changeDate, nextDueDate, dueDate,
      meetingTime, note, progress, by
    },
    ...
  }
  ```

### Changelog

Per-day changelog per sheet, stored in Document Properties:

* `KANBAN_CHANGELOG_S<sheetId>` = JSON array of entries:

  ```js
  {
    ts,        // timestamp ms
    row,       // row number at time of edit
    uid,       // row UID
    field,     // field key (e.g. 'note')
    from,      // previous value
    to,        // new value
    task,      // snapshot of task name at that time
    by,        // snapshot of By
    who        // editor email
  }
  ```

Used by:

* `getChangeLogData(opts)` – for UI.
* `getChangeLogCsv(opts)` – export.
* Nightly preview & realtime compare.

---

## Troubleshooting & FAQ

### Q1: Messages aren’t arriving in LINE

* Check Script Properties:

  * `OA_CHANNEL_ACCESS_TOKEN` is set and valid.
  * `OA_GROUP_ID` is set correctly (starts with `C` for group).
* Check execution logs:

  * Look for errors in `sendLineOA_`.
* Try `testSend()`:

  * Should push a test message to your group.

### Q2: Nothing happens when I edit the sheet

* Ensure you have **an installable `onEdit` trigger**:

  * Go to **Apps Script → Triggers**.
  * Verify a trigger for function `onEdit` is present and bound to the spreadsheet.
  * If not, run `createRealtimeEditTrigger()` manually.
* Make sure the edited sheet is in `CONFIG.SHEETS`.
* Ensure `HEADER_ROW` matches the actual header row (default 2).
* Check “Edit lock”:

  * If active and you’re outside window, edits will revert.

### Q3: Nightly / morning messages don’t show anything

* Is there any row:

  * With `Acted Date` equal to today’s custom day? (for nightly)
  * With `Next Due Date == today`? (for morning)
* Is `CONFIG.SEND_IF_EMPTY` set to `false`?

  * In that case no message is sent when there are no matching rows.
* Run preview functions:

  * `previewNightlyAlertMessage()`
  * `previewDueListMessage()`

### Q4: I get “Header ... not found” errors

* Check that headers in the Sheet **exactly match** `CONFIG.HEADERS`.
* Beware of:

  * Extra spaces
  * Non-breaking spaces
  * Different capitalization
* You can also move the header row and adjust `CONFIG.HEADER_ROW`.

### Q5: I changed Morning/Night times – why “today” feels wrong?

* The **custom day boundary** uses Morning hour.
* When you change Morning time, the script updates `CUSTOM_DAY_START_HOUR`.
* That affects how “today” is calculated for acted rows and snapshots.

---

## Contributing

Pull requests are welcome!

Ideas:

* Support multiple OA groups with different filters.
* Add per-user notification opt-in/opt-out.
* Add a small “Logs” sheet to record errors or send summaries.
* Localization / multi-language support for message texts.

When contributing:

1. Fork the repository.
2. Create a feature branch.
3. Make changes & add comments where logic is complex.
4. Open a Pull Request with a clear description.

---
