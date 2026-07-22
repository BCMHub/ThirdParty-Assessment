# Third-Party Assessment Tracker — Implementation Plan (v2)

**Status:** DRAFT v2 for adversarial review round 2 (Fable ⇄ Sol). Not yet built.
**Author:** Fable 5 (planner). **Builder:** GPT‑5.6 Sol.
**Changes from v1:** incorporates Sol round‑1 critique + user decisions:
Windows-identity auth (no passwords), normalized Vendor↔Assessment model,
derived status/dates via a single transition function, draft-first Outlook
email with an immutable email-event log, Excel-native write-lock concurrency,
import hardening, and a reproducible Windows build with a self-test.

---

## 1. Purpose

A self-contained Excel macro-enabled workbook (`.xlsm`) on a shared network
folder that tracks the lifecycle of third-party (vendor) assessments. A
Third-Party Assessor adds vendors (individually or by CSV/Excel import),
prepares Outlook emails to the internal **Business Owner** who drives each
assessment, logs meetings, and records assessment status manually. A **Manager**
sees everything, manages the user allowlist and settings. A dashboard summarizes
status. Email is prepared locally through **Classic** Microsoft Outlook via
COM automation — **no external API, no server, no cloud.**

## 2. Confirmed constraints & decisions

| Item | Decision |
|---|---|
| Platform | Windows + **Classic** Outlook desktop (New Outlook unsupported) |
| Host app | Excel macro workbook (`.xlsm`) |
| Sharing | Single `.xlsm` on a shared network folder, **single-writer** model |
| Identity | **Windows identity** (`USERNAME`) mapped to a role via allowlist — no passwords stored |
| Roles | Third-Party Assessor, Manager |
| Email transport | Classic Outlook COM; **draft-first** (default) |
| Bulk add | Import from Excel/CSV, validated, insert/update modes |
| Recipient | Internal **Business Owner** email |
| Assessment status | Not Started → In Progress → Submitted → Completed (derived invariants) |
| Vendor↔assessment | **One vendor → many assessments** (by cycle/year) |
| Follow-up "due" | ≥ N days since last email event AND status not Submitted/Completed |
| N (threshold) | Default 7, configurable in Settings |
| Dashboard | KPI tiles + status chart + overdue/due drill-down |
| Deliverable | VBA **source** + one-click Windows build script + manual fallback guide |

## 3. Data model (normalized)

### `tblVendors` (very hidden ListObject) — one row per vendor
| Column | Type | Notes |
|---|---|---|
| VendorID | Long PK | surrogate, from Settings counter |
| VendorName | Text | required; **not unique** (warn on dup, don't block) |
| BusinessOwnerName | Text | required |
| BusinessOwnerEmail | Text | required, format-validated — the email recipient |
| VendorContactPerson | Text | |
| VendorContactEmail | Text | validated if present |
| VendorContactPhone | Text (string) | preserve leading 0/+ |
| Notes | Text | |
| CreatedBy/CreatedOn, ModifiedBy/ModifiedOn | | audit stamps |

### `tblAssessments` (very hidden) — one row per assessment cycle
| Column | Type | Notes |
|---|---|---|
| AssessmentID | Long PK | **the row key** |
| VendorID | Long FK → tblVendors | |
| Cycle | Text | e.g. "2026" or "2026-H1" (label, user-defined) |
| Status | Enum | Not Started / In Progress / Submitted / Completed |
| SubmissionDate | Date | set by transition into Submitted (manual) |
| CompletedDate | Date | set by transition into Completed |
| MeetingsConducted | Long | incremented via "Log Meeting" |
| Notes | Text | |
| CreatedBy/On, ModifiedBy/On | | |

**Derived, never stored raw:**
- `ProcessStarted` = (StatusRank(Status) ≥ StatusRank("In Progress"))
- `EmailSentDate` = min(SentOn) over Initial email events for the assessment
- `LastFollowupDate` = max(SentOn) over Follow-up email events
- `FollowupCount` = count of Follow-up events in state Sent
- `NextFollowupDue` = (last email SentOn) + N, blank if Submitted/Completed or no email yet

### `tblEmailEvents` (very hidden, append-mostly, immutable log)
| Column | Type | Notes |
|---|---|---|
| EventID | Long PK | |
| AssessmentID | Long FK | |
| Kind | Enum | Initial / Followup |
| State | Enum | Prepared → DraftCreated → Sent / Failed / Cancelled |
| Recipient | Text | resolved BusinessOwnerEmail at prepare time |
| Subject | Text | |
| SendingAccount | Text | Outlook account SMTP used |
| DraftEntryID | Text | Outlook `EntryID` of created draft (for reconciliation) |
| PreparedOn / SentOn | Date | |
| PreparedBy | Text | Windows user |

### `tblUsers` (very hidden) — role allowlist (NO passwords)
| Column | Type | Notes |
|---|---|---|
| WindowsUsername | Text (PK) | matched case-insensitively to `Environ("USERNAME")` |
| DisplayName | Text | |
| Role | Enum | Assessor / Manager |
| Active | Bool | |
| CreatedBy/On | | |

First-run bootstrap: if `tblUsers` is empty, the current Windows user is seeded
as **Manager** and prompted to add others.

### `tblAudit` (very hidden, append-only)
Timestamp, WindowsUsername, Action, EntityType, EntityID, Detail.

### `Settings` (very hidden)
FollowupDays (N, default 7), OrgName, EmailInitialTemplate, EmailFollowupTemplate,
SendingAccountSMTP (optional; blank = Outlook default), EmailMode
(Draft/Direct, default Draft), BatchCap (default 50), InternalDomainAllowlist,
NextVendorID, NextAssessmentID, NextEventID, SchemaVersion, StaleBackupKeep.

## 4. Status transition function (single source of truth)

`modStatus.ApplyTransition(assessmentID, newStatus)`:
- StatusRank map: Not Started=0, In Progress=1, Submitted=2, Completed=3.
- Allowed transitions (forward by any steps allowed; backward allowed for
  correction **Manager-only**, with an audit entry).
- On entering **Submitted**: require/prompt SubmissionDate (default today);
  clear CompletedDate.
- On entering **Completed**: require CompletedDate (default today); ensure
  SubmissionDate present (if missing, prompt).
- On leaving Submitted/Completed backward: clear the corresponding date.
- Every change → audit log; all reads of ProcessStarted/NextFollowupDue are
  computed, never compared as text with `>=`.

## 5. Email flow (draft-first, honest semantics)

1. **Prepare** (per-assessment button OR batch): resolve recipient
   (BusinessOwnerEmail), enforce internal-domain allowlist, render template
   (`{VendorName} {OwnerName} {OrgName} {Cycle}`), **fail on unresolved
   `{...}` tokens**. Write a `tblEmailEvents` row (State=Prepared) **before**
   touching Outlook (intent persisted first).
2. Create Outlook `MailItem`, set `SendUsingAccount` explicitly (Settings or
   default), `.Save` → lands in **Drafts**; capture `EntryID`; State=DraftCreated.
   Default mode never calls `.Send` from the macro.
3. User reviews drafts in Outlook and sends. On next app open (and via a
   "Reconcile sent" button), `modEmail.Reconcile` looks up each DraftCreated
   event's `EntryID`: if the item is now in **Sent Items** → State=Sent, stamp
   `SentOn`; if still in Drafts → unchanged; if deleted without send → offer
   Cancelled. `.Send` is treated as *queued*, never as delivery proof.
4. **Direct mode** (Settings, Manager opt-in): after showing recipient list +
   BatchCap + explicit confirm, calls `.Send` with `SendUsingAccount`; State=Sent
   optimistically, still reconciled against Sent Items. Handles Object Model
   Guard prompts gracefully (documented; may prompt per mail on unhardened PCs).
5. **Batch "prepare all due"**: selects assessments where `NextFollowupDue <= today`
   (plus never-emailed if user opts in), groups by BusinessOwner, respects
   BatchCap, shows a preview, creates drafts, reports per-item results,
   stop-on-error option.

## 6. Concurrency (shared folder, single-writer)

- Rely on **Excel's native write lock**. On open, if `ThisWorkbook.ReadOnly` is
  True → run in **Browse mode**: Dashboard + read-only vendor/assessment views
  work; every mutating action is blocked with "Opened read-only — another user
  is editing. Try later." No custom lock file, no age-based reclaim (both were
  race-prone).
- On open (when writable), write a **timestamped backup copy** to a `Backups\`
  subfolder (keep last `StaleBackupKeep`, default 10) for crash recovery.
- Documented operational reality: this is a **one-writer-at-a-time** tool; if
  someone leaves it open for edit, others are read-only until they close.

## 7. Identity & roles

- `modIdentity.CurrentUser()` = `Environ("USERNAME")` (fallback
  `CreateObject("WScript.Network").UserName`), matched case-insensitively to
  `tblUsers`. Unknown/inactive user → Browse-only, with a "request access"
  message naming the Managers.
- `RequireRole("Manager")` guards Manager-only actions; Assessor-only buttons
  are also hidden for Assessors (defense in depth). VBA is **not** a security
  boundary — stated plainly in README; real protection is share ACLs.

| Capability | Assessor | Manager |
|---|---|---|
| View dashboard/vendors/assessments | ✅ | ✅ |
| Add/edit vendor & assessment | ✅ | ✅ |
| Import | ✅ | ✅ |
| Prepare emails / follow-ups | ✅ | ✅ |
| Log meeting, set status forward | ✅ | ✅ |
| Backward status correction | ❌ | ✅ |
| Delete vendor/assessment | ❌ | ✅ |
| Manage users, Settings, Direct-send mode | ❌ | ✅ |

## 8. Import (hardened)

- `Application.GetOpenFilename`; accept `.csv`/`.xlsx` only; open imported files
  with links/macros disabled; never trust formulas.
- **Formula-injection neutralization**: any imported cell value beginning with
  `= + - @ TAB CR` is prefixed with `'` before storage/display.
- Two modes: **Insert** (new vendors/assessments) and **Update** (match on a
  business key = VendorName + BusinessOwnerEmail, case-insensitive; user picks
  which fields update).
- CSV: explicit UTF-8 handling and a documented delimiter/locale note.
- Validate every row (required fields, email format); stage in `ImportStaging`;
  show `frmImportReview` ("N valid, M errors" + per-row reasons); commit only
  valid rows on confirm; audit the import.

## 9. Environment self-check (startup)

`modEnv.Check` verifies and reports clearly on failure:
- Macros enabled (WELCOME sheet fallback text if not).
- Classic Outlook reachable (`CreateObject("Outlook.Application")`) and **not**
  New-Outlook-only; a configured MAPI profile exists.
- Workbook not opened from a browser temp/quarantine path.
- Write access to folder + `Backups\`.
- SchemaVersion matches; offer migration if older.
- Excel bitness noted for support.

## 10. Modules & forms

**Modules:** `modIdentity`, `modVendors`, `modAssessments`, `modStatus`,
`modEmail` (prepare + reconcile), `modFollowup`, `modImport`, `modDashboard`,
`modUsers`, `modSettings`, `modEnv`, `modBackup`, `modAudit`, `modUtil`
(StatusRank, email validation, formula-injection guard, ID allocation).
**Forms:** `frmVendor`, `frmAssessment`, `frmStatus`, `frmImportReview`,
`frmUsers`, `frmSettings`, `frmBatchPreview`.

## 11. Deliverables (reproducible on the user's Windows PC)

The dev host here is macOS with no Excel, so the compiled `.xlsm` is assembled
on the user's Windows machine:
1. `/src/*.bas`, `/src/*.cls`, `/src/*.frm` — all VBA source.
2. `Build_Windows.ps1` (one-click): creates a new `.xlsm`, builds sheets/tables/
   named ranges/Settings, imports all modules via the VBIDE object model, runs
   `modEnv.SelfTest`, saves. **Requires** "Trust access to the VBA project object
   model" (guide shows how to toggle). Prints PASS/FAIL of self-test.
3. `ASSEMBLY_GUIDE.md` — manual fallback (paste modules into a blank macro
   workbook) for locked-down PCs where VBOM trust is disallowed.
4. `README.md` — IT prerequisites (Classic Outlook, macro policy, code-signing
   the VBA project with a trusted cert, share ACLs, Trusted Location caveats,
   32/64-bit note), first-run bootstrap, backup/recovery, security posture.
5. `sample_vendors.csv` — import template with correct headers.
6. `modEnv.SelfTest` (built-in) + `ACCEPTANCE_TESTS.md` — a Windows test
   checklist (offline/cached Outlook, read-only open, import edge cases,
   status transitions with date invariants, batch cap, reconcile, backup/restore,
   ~10k-row performance) since Windows runtime testing can't happen on macOS.

## 12. Non-goals / documented limitations

- Not a security boundary; not real multi-writer concurrency; New Outlook
  unsupported; `.Send` is queue-not-delivery; requires IT enablement of macros.
  All stated in README so expectations are set.

## 14. Correctness clauses — AUTHORITATIVE (Sol round 2). These override any earlier conflicting text.

### Identity
- Current user = normalized `USERDOMAIN\USERNAME` (fallback `WScript.Network`
  UserDomain+UserName). Bare `USERNAME` is not unique. If identity cannot be
  resolved → **deny all mutations** (Browse mode).

### Users bootstrap fails closed
- **The build script provisions the initial Manager** (the Windows identity of
  whoever runs the build, or a value passed to it) and stamps a `Bootstrapped`
  marker. The running app **never** auto-grants Manager.
- Workbook state is explicit: `BootstrapPending` (build-created, awaiting first
  Manager) vs `Bootstrapped`. Any **missing, unknown, or corrupt** state, or an
  empty/corrupt `tblUsers` in a `Bootstrapped` workbook → **deny all access**
  (Browse-only + alert). Fail closed, never fail open.

### Assessment owns its business-owner snapshot
- `tblAssessments` **snapshots** `BusinessOwnerName` + `BusinessOwnerEmail` at
  creation. `tblVendors` holds the *current default* owner. Editing the vendor's
  owner does **not** retroactively change past cycles. Email recipient resolves
  from the assessment's snapshot.
- `(VendorID, Cycle)` is **unique** — reject duplicates.

### Referential integrity — soft delete
- No hard delete when dependents exist. Vendor/assessment delete = **archive/
  soft-delete** (`Archived` flag + timestamp + who), hidden from active views,
  preserved for events/audit. Hard delete allowed **only** when zero dependent
  assessments/events, Manager-only, audited.
- Every append/update runs an explicit **FK check** first (assessment→vendor,
  event→assessment); orphan writes are rejected.

### ID allocation (crash/reentrancy safe under single-writer)
- Settings store `NextVendorID` / `NextAssessmentID` / `NextEventID` (the *next*
  id to hand out). Allocate `newID = max(NextStoredID, MAX(existing IDs) + 1)`;
  persist `NextStoredID = newID + 1`; assert the id is unused before write.
- This guarantees **monotonic** allocation (ids only increase, never collide),
  **not** gap-free allocation — a crash between allocate and commit may leave a
  skipped id, which is acceptable and never reused. (True gap-free allocation is
  impossible without transactional storage.)
- Disable double-submit on all forms/buttons (re-entrancy guard) so a slow save
  can't mint two rows.

### Email events = mutable lifecycle record with audited transitions
- Not claimed immutable. States: `Prepared → DraftCreated → Queued →
  Sent | Unresolved | Failed | Cancelled`. Every state change is audited.
- **GUID correlation token** is generated and the `tblEmailEvents` row written
  (State=Prepared) **before** any Outlook call, so a crash mid-create leaves a
  Prepared row to reconcile — no untracked draft, no duplicate on retry.
- On draft creation: stamp the GUID as a custom Outlook `UserProperty` on the
  MailItem, and store the draft `EntryID` **and `StoreID`** (EntryID alone is
  unstable across folder moves). State=DraftCreated.
- Never write `Sent` optimistically after `.Send` — direct-send sets `Queued`.
- **Reconciliation** searches the correct account/shared mailbox **by GUID
  UserProperty** (not EntryID) across **Drafts, Outbox, and Sent Items**, and
  maps folder → state:
  - found in **Drafts** → stays `DraftCreated`
  - found in **Outbox** → `Queued`
  - found in **Sent Items** → `Sent` + stamp `SentOn` (exactly 1 match; >1 →
    **manual review**)
  - found **nowhere** after a defined grace period → `Unresolved` (accounts for
    disabled Sent-copy saving, moved/archived items — surfaced, never silently
    "deleted"); a later pass may still resolve `Unresolved → Sent`.
  The state graph explicitly permits `DraftCreated → Sent` (user sent the draft
  manually). Capture `PR_INTERNET_MESSAGE_ID` after send only as secondary
  evidence.
- Suppress creating a new active (`Prepared`/`DraftCreated`/`Queued`) event when
  one already exists for that assessment+kind — no duplicate drafts.
- **Follow-up is blocked until the Initial event is confirmed `Sent`.**

### Status invariants (enforced by the single `ApplyTransition`)
- `Not Started` / `In Progress`: both SubmissionDate & CompletedDate blank.
- `Submitted`: SubmissionDate present, CompletedDate blank.
- `Completed`: both present, `SubmissionDate <= CompletedDate <= today`.
- Same status with **no field change** = no-op. Same status **with a date change**
  is a distinct **Manager-only correction path** routed through `ApplyTransition`
  (`mode = Correction`): it re-validates the target-status invariants against the
  new dates and writes row + audit **atomically**. Validate all inputs **before**
  mutating; if the user cancels or the audit write fails, the row is unchanged.
- Backward transitions & date corrections are **Manager-only**, audited, and go
  through `ApplyTransition` — nothing edits dates directly.
- Imports and forms **cannot bypass** `ApplyTransition`.
- Startup self-test scans for pre-existing invalid rows and **blocks mutation**
  until repaired.

### Import (text-first)
- Parse **CSV as raw text** (never open via Excel before neutralization).
- For `.xlsx`: open with links/events/macros/calculation disabled; **reject
  formula cells**. Neutralize any value starting `= + - @ TAB CR` with a `'`.
- Update-mode match key documented; changing owner does not silently duplicate.

### Internal-domain allowlist
- Match on the **exact domain after `@`** (boundary-aware), never substring.

### Backups
- **Once per day per workbook** (not per user). Force an extra backup immediately
  **before** schema migration, bulk import, or destructive recovery. A backup
  failure is **visible and blocks** those high-risk operations. (Same-share
  copies are recovery, not disaster recovery — documented.)

### Build script must handle
- `ThisWorkbook`/sheet-module code, `.frm` **plus `.frx`** resources, project
  **References**, sheet **CodeNames**, a schema manifest, and a **reopen +
  self-test** gate that fails the build on any error. The from-scratch VBIDE
  script is the source of truth; any distributed prebuilt `.xlsm` is an artifact,
  not a parallel source.

## 13. Round-2 questions for Sol (answered — see §14)

1. Is draft-first + Sent-Items reconciliation the right honest model, or is an
   optional per-message read-receipt/delivery-receipt worth adding?
2. Is `EntryID` stable enough for reconciliation, or should we also stamp a
   custom `UserProperty`/`PR_INTERNET_MESSAGE_ID` correlation token?
3. Backup-on-open of a shared `.xlsm`: acceptable, or too slow for large files —
   throttle to once/day per user?
4. Any remaining status/date invariant holes in §4?
5. Is the VBIDE-import build script the most reliable path, or should we ship a
   pre-seeded skeleton `.xlsm` (data sheets only) and import just the code?
