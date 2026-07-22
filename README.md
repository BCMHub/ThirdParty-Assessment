# Third-Party Assessment Tracker

This repository is the reproducible source of an Excel/VBA third-party assessment tracker. It intentionally does **not** contain a compiled `.xlsm`: run `Build_Windows.ps1` on a Windows PC with desktop Excel. Classic Outlook is a runtime requirement, not a build-machine requirement.

The design follows `PLAN.md`, with Section 14 taking precedence over earlier sections. Authoritative data is held in normalized, very-hidden Excel tables; visible sheets are projections and forms are the supported write path.

## Build on Windows

Prerequisites:

- Windows 10/11 with desktop Microsoft Excel. The structural-only build gates do not launch Outlook or require a configured MAPI profile.
- Windows PowerShell 5.1 or PowerShell 7 running in the interactive user's Windows session.
- Excel setting enabled: **File → Options → Trust Center → Trust Center Settings → Macro Settings → Trust access to the VBA project object model**. Close every Excel process after changing it. Group Policy may prevent this; use `ASSEMBLY_GUIDE.md` or ask IT.
- Macro execution permitted by organizational policy. Files received from the internet may carry Mark-of-the-Web; unblock the ZIP/source before building if IT policy allows.
- Permission to create the output workbook. Runtime users also need permission to create a `Backups` subfolder beside it.

From PowerShell in this directory:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\Build_Windows.ps1 -InternalDomain "contoso.com"
```

Useful options:

```powershell
.\Build_Windows.ps1 `
  -OutputPath "S:\TPA\ThirdParty_Assessment_Tracker.xlsm" `
  -InitialManager "CONTOSO\jsmith" `
  -InitialManagerDisplayName "Jordan Smith" `
  -InternalDomain "contoso.com;subsidiary.example" `
  -Force
```

`-InitialManager` defaults to the build runner's `USERDOMAIN\USERNAME`. `-Force` replaces only the exact `-OutputPath`; without it, an existing file causes a safe failure. If `-InternalDomain` and `USERDNSDOMAIN` are both unavailable, the builder uses the deliberately non-routable `example.invalid`, and a Manager must correct the allowlist before email can be prepared.

The builder creates sheets, ListObjects, named settings, the schema manifest, CodeNames, references, modules, forms, and document-module code. It starts at `BootstrapPending`, directly provisions the initial Manager, stamps `Bootstrapped`, runs `modEnv.SelfTest(False, True)` in structural-only mode, saves, reopens, and runs that structural gate again. These gates validate schema/tables/columns, CodeNames, users, data integrity, settings, and named ranges without Outlook, path/quarantine, read-only, backup-writability, mutation-block, or role-UI side effects. Any failed gate exits with `BUILD FAILED`; do not distribute that output.

## IT deployment checklist

- Place the signed, accepted workbook on a restricted Windows file share. Grant Modify to authorized operators and deny access to people who must not read the data. VBA roles are usability controls, **not a security boundary**.
- Digitally sign the final VBA project with an organizational code-signing certificate. Re-sign after every source rebuild/change. Distribute the signing certificate chain through enterprise trust policy.
- Prefer a signed-publisher macro policy. A Trusted Location disables several Office trust prompts for every supported file in that location; use one only when its ACLs and change controls are strong. Never make a broad writable share a Trusted Location merely to suppress prompts.
- Test both Office bitnesses used by the organization. The only Windows API declaration is conditional `VBA7`/`PtrSafe` GUID generation and is compatible with 32-bit and 64-bit Office.
- Keep Classic Outlook patched and configured. Cached/offline mode may queue mail; `.Send` is never treated as proof of delivery.
- Back up the share outside Excel. The workbook's same-share `Backups` copies are for quick recovery, not ransomware, server-loss, or disaster recovery.

## First run and identity

Identity is normalized `USERDOMAIN\USERNAME`, with `WScript.Network` as a fallback for both domain and user. A bare username is never accepted as unique identity.

The runtime never auto-creates a Manager. It fails closed when identity is unresolved, `BootstrapState` is missing/unknown/corrupt, the state is `BootstrapPending`, or a bootstrapped `tblUsers` is empty/corrupt/lacks an active Manager. Unknown or inactive users get Browse mode and a Manager contact list.

The initial Manager should immediately:

1. Open **Settings** and replace `OrgName`, confirm the exact semicolon-separated `InternalDomainAllowlist`, templates, follow-up days, sending account, and batch cap.
2. Open **Users** and add each user as `Assessor` or `Manager` using `DOMAIN\USERNAME`.
3. Confirm Draft mode. Enable Direct mode only after an Outlook/MAPI pilot and policy approval.
4. Run the Windows acceptance checklist in `ACCEPTANCE_TESTS.md`, then sign and distribute the accepted workbook.

## Operating model

- Excel's native file lock is the concurrency control. One user edits; later openers get read-only Browse mode. There is no custom lock file and no multi-writer merge.
- Assessments own a snapshot of business owner name/email. Changing a vendor's current default does not rewrite past assessment cycles.
- `(VendorID, Cycle)` is unique. IDs are monotonic and gap-tolerant: each allocation is `max(stored next, current max + 1)`, then the next counter is persisted before the row write.
- Vendors/assessments with dependents are archived, not deleted. Hard delete is Manager-only, audited, backed up first, and allowed only with zero dependents.
- All status/date changes go through `modStatus.ApplyTransition`. Backward moves and same-status date corrections are Manager-only. Validation happens before writes; an audit failure restores the original row.
- Every email intent gets a persisted GUID event before Outlook is called. The GUID is stamped into an Outlook custom user property. Drafts, Outbox, and Sent Items are reconciled by GUID and store/account; `EntryID` and internet message ID are secondary evidence.
- Default email behavior creates a draft. Direct mode records `Queued`, never optimistic `Sent`. Only an exact single GUID match in Sent Items becomes `Sent`.
- Follow-up is unavailable until an Initial event is confirmed Sent. Due dates are derived from confirmed sent events and current settings; they are not stored source fields.
- Email domain checks compare the exact text after `@` with the allowlist. `evilcontoso.com` does not match `contoso.com`.

## Import

Use `sample_vendors.csv` as the header template. CSV is decoded as UTF-8 raw text and parsed before values reach Excel cells. `.xlsx` is opened read-only with links, events, macros, and automatic calculation disabled; any formula cell is rejected.

Insert mode creates vendors and optional assessments. Update mode matches an existing active vendor by case-insensitive `VendorName + BusinessOwnerEmail`; a missing match is an error and never silently becomes an insert. The review form lets the user select which non-key fields to update. Key fields stay unchanged, so changing the owner email is a separate in-app vendor edit, not an ambiguous import update. An optional `Cycle` creates a new assessment only when `(VendorID, Cycle)` does not already exist.

Imported values beginning with `=`, `+`, `-`, `@`, TAB, or CR are prefixed with an apostrophe before staging/storage. CSV uses a comma delimiter regardless of the Windows regional list separator; quote fields that contain commas, quotes, or line breaks.

## Backup and recovery

On a writable open, the app creates at most one normal backup per calendar day per workbook in `Backups\`, retaining `StaleBackupKeep` files. A forced extra backup is required immediately before bulk import and hard delete; failure is visible and blocks the operation.

Recovery:

1. Make all users close the live workbook.
2. Preserve the damaged file and copy the chosen backup to a separate recovery filename.
3. Open the recovery copy locally/read-only first and run `modEnv.SelfTest` from the Macro dialog.
4. Reconcile email events before retrying any email. Never create a replacement event while a prior event is `Prepared`, `DraftCreated`, or `Queued`.
5. After validation, have IT replace the shared copy with appropriate ACLs and re-sign if the file changed.

## Honest limitations

- Excel/VBA project protection and hidden sheets can be bypassed by a determined user with file access. Share ACLs, trusted signing, endpoint controls, and audit review provide the real controls.
- The workbook is single-writer and has no transactional database. Audit-backed rollback covers critical row operations, but Excel or Windows can still crash between saves. Monotonic IDs can have gaps by design.
- Outlook reconciliation searches Drafts, Outbox, and Sent Items in the captured/sending stores. Items moved to arbitrary archive folders, disabled Sent-copy settings, transport rules, or duplicate copied items may remain `Unresolved` and require manual review.
- Sent Items proves Outlook recorded a send, not recipient delivery, reading, or business completion. Cached/offline Outlook can delay transition from Outbox to Sent.
- Same-share backups do not protect against loss of the share. Use enterprise backup/versioning.
- The macOS source host cannot compile or execute Excel VBA. The Windows builder's two structural-only self-test gates and `ACCEPTANCE_TESTS.md` are required release evidence; startup/manual runtime self-tests additionally verify Outlook, path, write mode, and backup-folder access.

## Source layout

- `src/mod*.bas`: application modules.
- `src/frm*.frm` plus matching `.frx`: UserForms. The `.frx` files are explicit no-binary-resource placeholders because the forms use only textual control definitions.
- `src/ThisWorkbook.cls` and `src/sht*.cls`: document-module source installed into existing workbook/sheet components by the builder.
- `Build_Windows.ps1`: authoritative assembly automation.
- `ASSEMBLY_GUIDE.md`: locked-down/manual fallback.
- `ACCEPTANCE_TESTS.md`: Windows release gate.
