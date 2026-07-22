# Manual Assembly Guide

Use this fallback only when VBIDE automation is prohibited. The preferred and authoritative assembly path is `Build_Windows.ps1`; manual assembly is more error-prone and must still pass every acceptance test.

## 1. Create and save the container

1. On Windows desktop Excel, create a blank workbook and immediately save it as `ThirdParty_Assessment_Tracker.xlsm` (Excel Macro-Enabled Workbook).
2. Keep exactly these worksheets, with these visible names and VBA CodeNames:

| Sheet name | CodeName | Initial visibility |
|---|---|---|
| WELCOME | `shtWelcome` | Visible |
| Dashboard | `shtDashboard` | Visible |
| Vendors | `shtVendors` | Visible |
| Assessments | `shtAssessments` | Visible |
| ImportStaging | `shtImportStaging` | Hidden |
| Data_Vendors | `shtDataVendors` | VeryHidden |
| Data_Assessments | `shtDataAssessments` | VeryHidden |
| Data_EmailEvents | `shtDataEmailEvents` | VeryHidden |
| Data_Users | `shtDataUsers` | VeryHidden |
| Data_Audit | `shtDataAudit` | VeryHidden |
| Settings | `shtSettings` | VeryHidden |
| Schema | `shtSchema` | VeryHidden |

Set CodeNames in the VBE Properties window, not by renaming the sheet tab. Leave WELCOME visible until macros and self-test pass.

## 2. Create ListObjects exactly

Enter each header row at `A1`, select the header plus one blank row, choose **Insert → Table**, confirm headers, assign the exact table name under Table Design, then delete the blank data row.

| Table / sheet | Headers in exact order |
|---|---|
| `tblVendors` / Data_Vendors | `VendorID, VendorName, BusinessOwnerName, BusinessOwnerEmail, VendorContactPerson, VendorContactEmail, VendorContactPhone, Notes, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn, Archived, ArchivedOn, ArchivedBy` |
| `tblAssessments` / Data_Assessments | `AssessmentID, VendorID, Cycle, BusinessOwnerName, BusinessOwnerEmail, Status, SubmissionDate, CompletedDate, MeetingsConducted, Notes, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn, Archived, ArchivedOn, ArchivedBy` |
| `tblEmailEvents` / Data_EmailEvents | `EventID, AssessmentID, Kind, State, CorrelationToken, Recipient, Subject, SendingAccount, DraftEntryID, StoreID, InternetMessageID, PreparedOn, SentOn, PreparedBy, LastStateOn, ErrorDetail` |
| `tblUsers` / Data_Users | `WindowsUsername, DisplayName, Role, Active, CreatedBy, CreatedOn` |
| `tblAudit` / Data_Audit | `Timestamp, WindowsUsername, Action, EntityType, EntityID, Detail` |
| `tblSettings` / Settings | `SettingName, SettingValue, Description` |
| `tblImportStaging` / ImportStaging | `SourceRow, ImportMode, Action, Valid, Errors, VendorName, BusinessOwnerName, BusinessOwnerEmail, VendorContactPerson, VendorContactEmail, VendorContactPhone, Notes, Cycle` |
| `tblSchemaManifest` / Schema | `ObjectType, ObjectName, SheetName, Columns, SchemaVersion` |

Populate `tblSettings` with the same names/defaults shown in `Build_Windows.ps1` under `$settingRows`. Create workbook names `cfg<SettingName>` pointing to each corresponding `SettingValue` cell. Set `SchemaVersion` to `2.0.0`. Start `BootstrapState` as `BootstrapPending`.

Populate `tblSchemaManifest` with one row per ListObject (including pipe-separated headers) and one row per worksheet CodeName, all at schema version `2.0.0`. The exact automated construction in `Build_Windows.ps1` is the reference.

## 3. Provision bootstrap without running the app

Before importing/running application code, add exactly one `tblUsers` row:

- `WindowsUsername`: current Manager as uppercase `DOMAIN\USERNAME`
- `DisplayName`: human-readable name
- `Role`: `Manager`
- `Active`: `TRUE`
- `CreatedBy`: `MANUAL_ASSEMBLY`
- `CreatedOn`: current date/time

Add a `tblAudit` row with action `BOOTSTRAP`, then change `BootstrapState` from `BootstrapPending` to `Bootstrapped`. Do not use a runtime macro to grant the first Manager. If assembly is interrupted before all three steps are complete, leave `BootstrapPending`, discard the partial copy, and restart from a clean file.

## 4. Import VBA

In the VBE (`Alt+F11`):

1. Confirm references under **Tools → References**: Visual Basic for Applications, Microsoft Excel Object Library, OLE Automation, and Microsoft Office Object Library. Outlook is deliberately late-bound and does not need a project reference.
2. Import every `src/mod*.bas` using **File → Import File**.
3. Import every `src/frm*.frm`. Keep its matching `.frx` in the same directory. These forms have no binary assets, so the supplied `.frx` files are textual placeholders and are not referenced by the form streams.
4. Do **not** import `ThisWorkbook.cls` or `sht*.cls` as new class modules. Open each existing document module, delete its code, then copy from `Option Explicit` through EOF from the matching source file.
5. Choose **Debug → Compile VBAProject**. Correct any manual transcription/schema error before continuing; do not suppress it.

## 5. Build visible sheets and buttons

Follow the layout statements and `Add-Button` calls in `Build_Windows.ps1`. At minimum:

- WELCOME must explain how to enable approved macros when its fallback text remains visible.
- Dashboard needs KPI cells `B4`, `D4`, `F4`, `H4`; due-list headers in `A12:H12`; hidden status helper range `J2:K6`; and a chart named `chtStatus`.
- Vendors uses headers in `A4:H4`; Assessments uses `A4:N4`.
- Assign button macros exactly as in the build script. Names beginning `mut_` are hidden when mutation is blocked; names beginning `mgr_` are visible only to Managers.

## 6. Gate the workbook

1. Save and close Excel completely.
2. Reopen the `.xlsm` in Excel under a Windows profile with Classic Outlook/MAPI configured.
3. Run `modEnv.SelfTest` from `Alt+F8`. It must report PASS; this is the full runtime environment check.
4. Save, close, reopen, and run `modEnv.SelfTest` again. The automated builder instead passes `structuralOnly:=True` at both gates so its acceptance does not depend on Outlook, path/quarantine, read-only, or backup-folder state.
5. Execute all of `ACCEPTANCE_TESTS.md`, digitally sign the VBA project, and archive the exact accepted source/build log.

Any missing/corrupt bootstrap marker, user table, table header, FK, status invariant, GUID event token, Classic Outlook profile, or write access must leave the workbook in Browse mode. Do not “fix” a failure by weakening the self-test.
