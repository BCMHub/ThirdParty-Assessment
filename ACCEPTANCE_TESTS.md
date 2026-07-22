# Windows Acceptance Tests

Record Excel/Outlook versions, Office bitness, Windows build, tester identity, source commit/hash, output hash, test date, and evidence for every case. Run on a disposable acceptance copy and test mailbox before signing or deployment.

## A. Build and bootstrap gates

- [ ] Run `Build_Windows.ps1 -InternalDomain <real-domain>` with no output present. Confirm it provisions the current `USERDOMAIN\USERNAME` as Manager, prints both self-test PASS gates, and produces an `.xlsm` plus `Backups` when self-test checks write access.
- [ ] Disable **Trust access to the VBA project object model**, close Excel, and run the builder. Confirm it fails with the exact Trust Center remediation and does not claim success.
- [ ] Interrupt a build after `BootstrapPending` is written but before the Manager row/`Bootstrapped` stamp (use a disposable edited script or manual procedure). Open the partial workbook and confirm all mutations are denied; the app never self-grants Manager.
- [ ] In a copy, set `BootstrapState` blank, misspelled, and unknown. Separately remove all users, duplicate an identity, corrupt a role, and deactivate every Manager. Each case must fail closed and show Browse-only guidance.
- [ ] Run with `-InitialManager DOMAIN\user` and verify exact normalized identity. Try a bare username and confirm the builder rejects it.
- [ ] Confirm every worksheet CodeName, every required ListObject/header, named `cfg*` setting, schema-manifest row, standard module, UserForm, and document module exists after reopen.

## B. Identity, role, and write lock

- [ ] Open as active Assessor: view/add/edit/import/email/meeting/forward-status functions work; Users, Settings, hard delete, backward status, and date correction are denied or hidden.
- [ ] Open as active Manager: Manager controls are visible and guarded actions work.
- [ ] Open as unknown/inactive identity: visible views work, Manager contact names are shown, all mutations are denied.
- [ ] Open the shared file for edit as User A, then open it as User B. User B must see Excel read-only/Browse mode; every mutating entry point must display “Opened read-only - another user is editing. Try later.” No custom lock file should appear.

## C. Vendor, assessment, FK, archive, and IDs

- [ ] Add duplicate vendor names and confirm a warning allows the duplicate. Required fields and malformed optional contact email are rejected.
- [ ] Create two cycles for one vendor. Attempt the same `(VendorID, Cycle)` with different casing and confirm rejection.
- [ ] Change the vendor's default business owner. Confirm existing assessment owner snapshots do not change and new assessments use the new default.
- [ ] Attempt an assessment with nonexistent/archived VendorID through an immediate-window/API call; confirm the FK check rejects it.
- [ ] Archive a vendor/assessment with dependents and confirm all history/events remain. Confirm hard delete is rejected while dependents exist.
- [ ] Set each `Next*ID` counter below the maximum existing ID, then create a row. Confirm allocation is `MAX(existing)+1`, the counter becomes new ID + 1, no collision occurs, and prior gaps are not reused.
- [ ] Double-click Save/submit buttons or hold Enter during a slow share save. Confirm the re-entrancy guard creates one row/event only.

## D. Status/date invariants and audit atomicity

- [ ] Forward-transition by one and multiple steps. Verify: Not Started/In Progress have no dates; Submitted has SubmissionDate only; Completed has both dates and `SubmissionDate <= CompletedDate <= today`.
- [ ] Cancel every date prompt and confirm the assessment row remains byte-for-byte unchanged.
- [ ] As Assessor, attempt backward transition and same-status date correction; confirm denial.
- [ ] As Manager, perform backward transition and confirm corresponding dates clear. Perform same-status Correction and confirm it is separately audited.
- [ ] Attempt future dates, CompletedDate before SubmissionDate, or dates on Not Started/In Progress; confirm validation occurs before mutation.
- [ ] Simulate a failed correction audit write (protect `Data_Audit`, make its table unavailable in a disposable copy, or inject a controlled failure). Confirm `ApplyTransition` restores the complete original assessment row and reports that the audit write failed.
- [ ] Directly corrupt a status/date row, reopen, and confirm `modEnv.SelfTest` blocks all mutation until repaired.

## E. Outlook intent, drafts, sending, and reconciliation

- [ ] With Draft mode, prepare Initial. Confirm the event row with unique GUID and `Prepared` exists before Outlook work, the draft has `TPATCorrelationToken`, `SendUsingAccount` is set, and the event becomes `DraftCreated` with EntryID and StoreID.
- [ ] Force Outlook creation failure after intent persistence. Confirm a tracked `Failed` (or recoverable `Prepared`) event exists and no untracked retry is created.
- [ ] Attempt a second active Initial/Followup for the same assessment+kind in `Prepared`, `DraftCreated`, or `Queued`; confirm suppression.
- [ ] Attempt Followup before Initial is confirmed Sent; confirm it is blocked.
- [ ] Move/send a draft normally. Reconcile and confirm GUID search maps Drafts→`DraftCreated`, Outbox→`Queued`, and exactly one Sent Items match→`Sent` with SentOn. Confirm `DraftCreated→Sent` works when a user manually sends the draft.
- [ ] Move an item so EntryID changes, then reconcile. Confirm the GUID—not EntryID—resolves it when it remains in Drafts/Outbox/Sent.
- [ ] Send from a non-default configured account/shared store and confirm reconciliation searches both captured draft StoreID and sending-account delivery store.
- [ ] Copy an item so two GUID matches exist. Confirm state becomes `Unresolved` with manual-review detail; it must not choose one silently.
- [ ] Remove/move the item outside searched folders or disable Sent-copy saving; after `ReconcileGraceHours`, confirm `Unresolved`. Later restore a single matching item to Sent Items and confirm `Unresolved→Sent`.
- [ ] Verify captured `PR_INTERNET_MESSAGE_ID` is secondary evidence only.
- [ ] Enable Direct mode as Manager. Confirm the preview lists recipients/cap, explicit confirmation is required, `.Send` produces `Queued` (never optimistic `Sent`), and only reconciliation confirms Sent.
- [ ] Test cached/offline Outlook. Confirm Classic Outlook/MAPI self-test passes with a configured profile, direct items stay `Queued` in Outbox while offline, and reconcile marks Sent only after sync. Test no profile and New-Outlook-only: self-test must fail and block mutation.
- [ ] Test an unhardened Object Model Guard environment and record prompts/failures; ensure errors are tracked and do not produce duplicate drafts.

## F. Domain and templates

- [ ] With allowlist `contoso.com;subsidiary.example`, accept `user@contoso.com` case-insensitively and reject `user@evilcontoso.com`, `user@contoso.com.evil`, malformed addresses, and blank domains.
- [ ] Verify assessment snapshot email, not the vendor's later default owner, is the recipient.
- [ ] Verify `{VendorName}`, `{OwnerName}`, `{OrgName}`, and `{Cycle}` render. Add an unknown `{Token}` and confirm preparation fails before Outlook.

## G. Import hardening

- [ ] Import `sample_vendors.csv` as UTF-8 CSV with quoted comma/quote/newline cases. Confirm comma is used regardless of regional list separator and all columns map correctly.
- [ ] Test CSV cells beginning with `=`, `+`, `-`, `@`, TAB, and CR in every free-text field. Confirm staged/stored values begin with an apostrophe and Excel never evaluates them. Include `=HYPERLINK(...)`, `+1+1`, `-2+3`, and `@SUM(...)`.
- [ ] Confirm CSV is never opened as an Excel workbook before neutralization.
- [ ] Import `.xlsx` values-only and confirm success. Include any formula cell and confirm the whole workbook is rejected with its cell address while links, events, macros, and automatic calculation remain disabled during read.
- [ ] Test missing/wrong headers, invalid emails, missing required values, empty file, unterminated CSV quote, and unsupported extension. Confirm staging errors are explicit and invalid rows are not committed.
- [ ] Insert mode: confirm new vendors and optional cycles. Update mode: confirm case-insensitive `VendorName + BusinessOwnerEmail` matching, missing match stays an error (never inserts), owner-email changes do not silently create a duplicate, selected fields update, and duplicate assessment cycles are rejected.
- [ ] Deny write access to `Backups`, then commit import. Confirm forced backup failure is visible and blocks all rows.

## H. Batch, dashboard, backup, and recovery

- [ ] Set `BatchCap=3` with at least five due items. Confirm preview is grouped/sorted by owner, lists only three, reports the cap, and creates at most three.
- [ ] Include never-emailed assessments and confirm they are Initial; confirmed/due records are Followup. Terminal Submitted/Completed records never appear.
- [ ] Confirm ProcessStarted, Initial sent date, last follow-up, follow-up count, and NextFollowupDue are derived from status/events/settings and change after reconciliation without raw derived storage.
- [ ] Confirm KPI tiles, status chart, and due/overdue list agree with source tables and hide archived rows.
- [ ] Open writable twice on the same day and confirm only one normal dated backup exists. Force bulk import/hard delete and confirm an extra timestamped backup. Reduce `StaleBackupKeep` and confirm only the newest configured count remains.
- [ ] Deny backup-folder permission and confirm daily failure is visible; bulk/hard-delete operations must block. Document that normal low-risk edits may continue only if organizational policy accepts that choice.
- [ ] Restore a backup under a new name, run self-test, reconcile email, and verify counters recover monotonically before replacing the shared copy.

## I. Performance and compatibility

- [ ] Populate approximately 10,000 vendors, 10,000 assessments, and representative email events in a disposable copy. Time open/self-test, dashboard refresh, vendor view, assessment view, due batch selection, CSV staging, and reconciliation. Record timings and memory for both 32-bit and 64-bit Office where supported.
- [ ] Confirm dashboard and assessment projections use the single-pass email summary cache (no per-assessment full event-table scan). Confirm bulk import suppresses per-row dashboard refresh and refreshes once at the end.
- [ ] Scroll/filter visible views; confirm headers, date formats, chart, and buttons remain usable. Confirm no formula errors, broken references, compile errors, or default blank sheets.
- [ ] Compile under supported 32-bit and 64-bit Office. Confirm conditional `VBA7`/`PtrSafe` GUID declarations compile and GUID creation succeeds.

Release only after all mandatory boxes pass or a documented risk owner accepts a specific exception.

