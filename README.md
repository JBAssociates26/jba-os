# JBA OS

Internal transaction-management system for Jim Bunn & Associates, built on
Google Apps Script. Agents create listings, work is routed through a
configurable workflow engine, marketing and the transaction coordinator (TC)
get notified at each step, and a native in-app checklist system replaces
paper/Google Forms checklists for Pre-Listing, Photography, and MLS
Submission.

This repo is the version-controlled mirror of the live Apps Script project.
**The code runs in Apps Script, not here** -- GitHub is source control and
history, synced to Apps Script via [`clasp`](https://github.com/google/clasp).

## Architecture

| File | Responsibility |
|---|---|
| `appsscript.json` | Project manifest (timezone, web app access settings) |
| `src/Code.gs` | Config (`JBA_OS`), auth (`requireUser_`), transaction CRUD, sequential workflow engine, notification queue |
| `src/WorkflowEngine.gs` | Routes the parallel Photography + MLS Submission checklists and advances the transaction once both are complete |
| `src/NativeChecklistEngine.gs` | Checklist template/response engine; Pre-Listing, Photography, and MLS Submission checklist definitions |
| `src/NewListingWizard.gs` | New listing intake, validation, options for the creation wizard |
| `src/NewListingNotifications.gs` | Branded HTML emails to the listing agent and marketing coordinator on listing creation |
| `src/TransactionWorkspace.gs` | Transaction detail data: property/seller/team info, workflow timeline, activity log, notes |
| `src/SharedUtilities.gs` | Cross-file helpers (currently `escapeHtml_`) |
| `src/Index.html` | Full frontend SPA -- transaction cards, new-listing wizard, transaction workspace, native checklist UI |

### Data model

Everything lives in one Google Sheet ("JBA OS Database"), created and
seeded by `setupJBAOSVersion1()` in `Code.gs`. Key sheets: `Users`,
`Transactions`, `Workflow Stages`, `Workflow Actions`, `Checklist Templates`,
`Checklist Responses`, `Parallel Action Status`, `Activity Log`,
`Notifications Queue`.

### Workflow model

Transactions move through **stages** (e.g. Pre-Listing -> Listing Secured ->
Live -> Under Contract -> Closing -> Closed), each with one or more
**actions**. Most actions are sequential. Photography and MLS Submission are
a **parallel group** -- either can be completed first; the transaction only
advances once both are done. `WorkflowEngine.gs` owns that routing logic.

## Setup

### First time connecting this repo to your live Apps Script project

1. Install clasp and log in with the Google account that owns the JBA OS
   Apps Script project:
   ```
   npm install -g @google/clasp
   clasp login
   ```
2. Copy `.clasp.json.example` to `.clasp.json` and fill in your real script
   ID (find it in the Apps Script editor: Project Settings -> Script ID).
   `.clasp.json` is gitignored on purpose -- it's specific to your machine
   and should never be committed.
   ```
   cp .clasp.json.example .clasp.json
   ```
3. Pull the current live code to confirm the connection works, then diff it
   against this repo before pushing anything:
   ```
   clasp pull
   ```

### Day-to-day

```
clasp push   # push local changes to the live Apps Script project
clasp pull   # pull live changes back down (e.g. edits made in the browser editor)
clasp open   # open the project in the Apps Script editor
```

## Known issues / open items

- `WorkflowEngine.gs` writes to `Current Assigned Role` and `Updated By`
  columns on the Transactions sheet that are not part of the schema
  `setupTransactions_()` creates in `Code.gs`. This fails silently today
  (guarded by a column-existence check) rather than erroring, but those
  fields are not actually being persisted. Add the columns to the sheet, or
  remove the writes, to resolve.
- No automated tests. `Code.gs` has manual test functions
  (`testJBAOSAccess`, `testTransactionWorkspace`, `testNativeChecklistEngine`)
  meant to be run from the Apps Script editor.
