[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot "ThirdParty_Assessment_Tracker.xlsm"),
    [string]$InitialManager = "",
    [string]$InitialManagerDisplayName = "",
    [string]$InternalDomain = "",
    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Get-NormalizedIdentity {
    param([string]$Requested)
    if (-not [string]::IsNullOrWhiteSpace($Requested)) {
        $candidate = $Requested.Trim().ToUpperInvariant()
    } else {
        $domain = [Environment]::GetEnvironmentVariable("USERDOMAIN")
        $user = [Environment]::GetEnvironmentVariable("USERNAME")
        if ([string]::IsNullOrWhiteSpace($domain) -or [string]::IsNullOrWhiteSpace($user)) {
            throw "Cannot resolve the build runner as USERDOMAIN\USERNAME. Pass -InitialManager 'DOMAIN\USERNAME'."
        }
        $candidate = ("{0}\{1}" -f $domain.Trim(), $user.Trim()).ToUpperInvariant()
    }
    if ($candidate -notmatch '^[^\\]+\\[^\\]+$') {
        throw "InitialManager must be a normalized DOMAIN\USERNAME identity."
    }
    return $candidate
}

function Add-Worksheet {
    param([object]$Workbook, [string]$Name)
    $sheet = $Workbook.Worksheets.Add([System.Type]::Missing, $Workbook.Worksheets.Item($Workbook.Worksheets.Count))
    $sheet.Name = $Name
    return $sheet
}

function Add-Table {
    param([object]$Sheet, [string]$Name, [string[]]$Headers)
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $Sheet.Cells.Item(1, $i + 1).Value2 = $Headers[$i]
        $Sheet.Cells.Item(2, $i + 1).Value2 = ""
    }
    $range = $Sheet.Range($Sheet.Cells.Item(1, 1), $Sheet.Cells.Item(2, $Headers.Count))
    $table = $Sheet.ListObjects.Add(1, $range, $null, 1)
    $table.Name = $Name
    $table.TableStyle = "TableStyleMedium2"
    if ($table.ListRows.Count -gt 0) { $table.ListRows.Item(1).Delete() }
    $Sheet.Rows.Item(1).Font.Bold = $true
    $Sheet.Rows.Item(1).WrapText = $true
    $Sheet.Rows.Item(1).RowHeight = 32
    $range.EntireColumn.AutoFit() | Out-Null
    return $table
}

function Add-TableRow {
    param([object]$Table, [object[]]$Values)
    $row = $Table.ListRows.Add()
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $row.Range.Cells.Item(1, $i + 1).Value2 = $Values[$i]
    }
    return $row
}

function Set-CodeName {
    param([object]$Project, [object]$Worksheet, [string]$CodeName)
    $component = $Project.VBComponents.Item($Worksheet.CodeName)
    $component.Name = $CodeName
}

function Get-DocumentCodeBody {
    param([string]$Path)
    $lines = [System.IO.File]::ReadAllLines($Path)
    $start = -1
    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ($lines[$i].Trim() -eq "Option Explicit") { $start = $i; break }
    }
    if ($start -lt 0) { throw "No Option Explicit marker found in document source: $Path" }
    return [string]::Join([Environment]::NewLine, $lines[$start..($lines.Length - 1)])
}

function Install-DocumentCode {
    param([object]$Project, [string]$ComponentName, [string]$SourcePath)
    $component = $Project.VBComponents.Item($ComponentName)
    $codeModule = $component.CodeModule
    if ($codeModule.CountOfLines -gt 0) { $codeModule.DeleteLines(1, $codeModule.CountOfLines) }
    $codeModule.AddFromString((Get-DocumentCodeBody -Path $SourcePath))
}

function Add-ReferenceIfMissing {
    param([object]$Project, [string]$Guid, [int]$Major, [int]$Minor, [string]$Label)
    foreach ($reference in $Project.References) {
        if ($reference.Guid -eq $Guid) { return }
    }
    try {
        [void]$Project.References.AddFromGuid($Guid, $Major, $Minor)
    } catch {
        throw "Could not add required VBA reference '$Label' ($Guid): $($_.Exception.Message)"
    }
}

function Add-Button {
    param([object]$Sheet, [string]$Name, [string]$Caption, [string]$Macro, [double]$Left, [double]$Top, [double]$Width = 118, [double]$Height = 28)
    $shape = $Sheet.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.TextFrame2.TextRange.Text = $Caption
    $shape.TextFrame2.TextRange.Font.Size = 10
    $shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    $shape.Fill.ForeColor.RGB = 7626838
    $shape.Line.Visible = 0
    $shape.OnAction = $Macro
}

$sourceDirectory = Join-Path $PSScriptRoot "src"
$requiredModules = @(
    "modIdentity", "modVendors", "modAssessments", "modStatus", "modEmail", "modFollowup", "modImport",
    "modDashboard", "modUsers", "modSettings", "modEnv", "modBackup", "modAudit", "modUtil"
)
$requiredForms = @("frmVendor", "frmAssessment", "frmStatus", "frmImportReview", "frmUsers", "frmSettings", "frmBatchPreview")
$documentSources = [ordered]@{
    "ThisWorkbook" = "ThisWorkbook.cls"
    "shtWelcome" = "shtWelcome.cls"
    "shtDashboard" = "shtDashboard.cls"
    "shtVendors" = "shtVendors.cls"
    "shtAssessments" = "shtAssessments.cls"
    "shtImportStaging" = "shtImportStaging.cls"
    "shtDataVendors" = "shtDataVendors.cls"
    "shtDataAssessments" = "shtDataAssessments.cls"
    "shtDataEmailEvents" = "shtDataEmailEvents.cls"
    "shtDataUsers" = "shtDataUsers.cls"
    "shtDataAudit" = "shtDataAudit.cls"
    "shtSettings" = "shtSettings.cls"
    "shtSchema" = "shtSchema.cls"
}

foreach ($module in $requiredModules) {
    $path = Join-Path $sourceDirectory ($module + ".bas")
    if (-not (Test-Path -LiteralPath $path)) { throw "Missing source module: $path" }
}
foreach ($form in $requiredForms) {
    foreach ($extension in @(".frm", ".frx")) {
        $path = Join-Path $sourceDirectory ($form + $extension)
        if (-not (Test-Path -LiteralPath $path)) { throw "Missing form resource: $path" }
    }
}
foreach ($entry in $documentSources.GetEnumerator()) {
    $path = Join-Path $sourceDirectory $entry.Value
    if (-not (Test-Path -LiteralPath $path)) { throw "Missing document-module source: $path" }
}

$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = [System.IO.Path]::GetDirectoryName($OutputPath)
if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory | Out-Null }
if (Test-Path -LiteralPath $OutputPath) {
    if (-not $Force) { throw "Output already exists: $OutputPath. Re-run with -Force to replace this exact file." }
    Remove-Item -LiteralPath $OutputPath -Force
}

$managerIdentity = Get-NormalizedIdentity -Requested $InitialManager
if ([string]::IsNullOrWhiteSpace($InitialManagerDisplayName)) {
    $InitialManagerDisplayName = $managerIdentity.Substring($managerIdentity.IndexOf('\') + 1)
}
if ([string]::IsNullOrWhiteSpace($InternalDomain)) {
    $InternalDomain = [Environment]::GetEnvironmentVariable("USERDNSDOMAIN")
}
if ([string]::IsNullOrWhiteSpace($InternalDomain)) {
    $InternalDomain = "example.invalid"
    Write-Warning "USERDNSDOMAIN was unavailable. InternalDomainAllowlist is set to example.invalid; a Manager must replace it before preparing email."
}
$InternalDomain = $InternalDomain.Trim().TrimStart([char]'@').ToLowerInvariant()

$excel = $null
$workbook = $null
$reopened = $null
$oldAutomationSecurity = $null
try {
    Write-Host "Starting Excel and creating workbook..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $oldAutomationSecurity = $excel.AutomationSecurity
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) { $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete() }
    $welcome = $workbook.Worksheets.Item(1)
    $welcome.Name = "WELCOME"
    $dashboard = Add-Worksheet $workbook "Dashboard"
    $vendorsView = Add-Worksheet $workbook "Vendors"
    $assessmentsView = Add-Worksheet $workbook "Assessments"
    $importStaging = Add-Worksheet $workbook "ImportStaging"
    $dataVendors = Add-Worksheet $workbook "Data_Vendors"
    $dataAssessments = Add-Worksheet $workbook "Data_Assessments"
    $dataEmailEvents = Add-Worksheet $workbook "Data_EmailEvents"
    $dataUsers = Add-Worksheet $workbook "Data_Users"
    $dataAudit = Add-Worksheet $workbook "Data_Audit"
    $settings = Add-Worksheet $workbook "Settings"
    $schema = Add-Worksheet $workbook "Schema"

    $workbook.SaveAs($OutputPath, 52)

    try {
        $project = $workbook.VBProject
        $null = $project.VBComponents.Count
    } catch {
        throw @"
Excel denied access to the VBA project object model.
Enable: File > Options > Trust Center > Trust Center Settings > Macro Settings >
'Trust access to the VBA project object model', close every Excel process, and re-run this script.
Group Policy may control this option; if so, ask IT or use ASSEMBLY_GUIDE.md.
Original error: $($_.Exception.Message)
"@
    }

    Set-CodeName $project $welcome "shtWelcome"
    Set-CodeName $project $dashboard "shtDashboard"
    Set-CodeName $project $vendorsView "shtVendors"
    Set-CodeName $project $assessmentsView "shtAssessments"
    Set-CodeName $project $importStaging "shtImportStaging"
    Set-CodeName $project $dataVendors "shtDataVendors"
    Set-CodeName $project $dataAssessments "shtDataAssessments"
    Set-CodeName $project $dataEmailEvents "shtDataEmailEvents"
    Set-CodeName $project $dataUsers "shtDataUsers"
    Set-CodeName $project $dataAudit "shtDataAudit"
    Set-CodeName $project $settings "shtSettings"
    Set-CodeName $project $schema "shtSchema"

    $tables = [ordered]@{}
    $tables["tblVendors"] = Add-Table $dataVendors "tblVendors" @("VendorID","VendorName","BusinessOwnerName","BusinessOwnerEmail","VendorContactPerson","VendorContactEmail","VendorContactPhone","Notes","CreatedBy","CreatedOn","ModifiedBy","ModifiedOn","Archived","ArchivedOn","ArchivedBy")
    $tables["tblAssessments"] = Add-Table $dataAssessments "tblAssessments" @("AssessmentID","VendorID","Cycle","BusinessOwnerName","BusinessOwnerEmail","Status","SubmissionDate","CompletedDate","MeetingsConducted","Notes","CreatedBy","CreatedOn","ModifiedBy","ModifiedOn","Archived","ArchivedOn","ArchivedBy")
    $tables["tblEmailEvents"] = Add-Table $dataEmailEvents "tblEmailEvents" @("EventID","AssessmentID","Kind","State","CorrelationToken","Recipient","Subject","SendingAccount","DraftEntryID","StoreID","InternetMessageID","PreparedOn","SentOn","PreparedBy","LastStateOn","ErrorDetail")
    $tables["tblUsers"] = Add-Table $dataUsers "tblUsers" @("WindowsUsername","DisplayName","Role","Active","CreatedBy","CreatedOn")
    $tables["tblAudit"] = Add-Table $dataAudit "tblAudit" @("Timestamp","WindowsUsername","Action","EntityType","EntityID","Detail")
    $tables["tblSettings"] = Add-Table $settings "tblSettings" @("SettingName","SettingValue","Description")
    $tables["tblImportStaging"] = Add-Table $importStaging "tblImportStaging" @("SourceRow","ImportMode","Action","Valid","Errors","VendorName","BusinessOwnerName","BusinessOwnerEmail","VendorContactPerson","VendorContactEmail","VendorContactPhone","Notes","Cycle")
    $tables["tblSchemaManifest"] = Add-Table $schema "tblSchemaManifest" @("ObjectType","ObjectName","SheetName","Columns","SchemaVersion")

    $settingRows = [ordered]@{
        "FollowupDays" = @("7", "Days after the last confirmed sent email before follow-up is due.")
        "OrgName" = @("Your Organization", "Organization name used in templates.")
        "EmailInitialSubject" = @("Third-Party Assessment: {VendorName} ({Cycle})", "Initial email subject template.")
        "EmailInitialTemplate" = @("Hello {OwnerName},`r`n`r`nPlease begin the {Cycle} assessment for {VendorName} on behalf of {OrgName}.`r`n`r`nThank you.", "Initial plain-text email body.")
        "EmailFollowupSubject" = @("Follow-up: {VendorName} assessment ({Cycle})", "Follow-up email subject template.")
        "EmailFollowupTemplate" = @("Hello {OwnerName},`r`n`r`nThis is a follow-up for the {Cycle} assessment of {VendorName} for {OrgName}.`r`n`r`nThank you.", "Follow-up plain-text email body.")
        "SendingAccountSMTP" = @("", "Exact Outlook account SMTP; blank uses the first configured account.")
        "EmailMode" = @("Draft", "Draft by default. Direct is Manager-only and records Queued until reconciliation.")
        "BatchCap" = @("50", "Maximum items in one batch.")
        "InternalDomainAllowlist" = @($InternalDomain, "Semicolon-separated exact domains after @; no substring matching.")
        "NextVendorID" = @("1", "Next stored monotonic vendor ID.")
        "NextAssessmentID" = @("1", "Next stored monotonic assessment ID.")
        "NextEventID" = @("1", "Next stored monotonic email event ID.")
        "SchemaVersion" = @("2.0.0", "Expected schema version.")
        "StaleBackupKeep" = @("10", "Number of same-share recovery copies to retain.")
        "ReconcileGraceHours" = @("24", "Hours before a missing Outlook item becomes Unresolved.")
        "BootstrapState" = @("BootstrapPending", "Build-only state marker; runtime never grants Manager.")
        "LastBackupDate" = @("", "Reserved support field; filenames enforce once-per-day backups.")
    }
    $settingRowMap = @{}
    foreach ($name in $settingRows.Keys) {
        $row = Add-TableRow $tables["tblSettings"] @($name, $settingRows[$name][0], $settingRows[$name][1])
        $settingRowMap[$name] = $row.Range.Row
    }

    $now = Get-Date
    [void](Add-TableRow $tables["tblUsers"] @($managerIdentity, $InitialManagerDisplayName, "Manager", $true, "BUILD_SCRIPT", $now))
    [void](Add-TableRow $tables["tblAudit"] @($now, $managerIdentity, "BOOTSTRAP", "Workbook", "InitialManager", "Build script provisioned initial Manager and changed BootstrapPending to Bootstrapped."))
    $settings.Cells.Item($settingRowMap["BootstrapState"], 2).Value2 = "Bootstrapped"

    foreach ($name in $settingRowMap.Keys) {
        $safeName = "cfg" + $name
        $formula = "='Settings'!`$B`$" + $settingRowMap[$name]
        [void]$workbook.Names.Add($safeName, $formula)
    }

    foreach ($tableName in $tables.Keys) {
        $table = $tables[$tableName]
        $columnNames = @()
        foreach ($column in $table.ListColumns) { $columnNames += $column.Name }
        [void](Add-TableRow $tables["tblSchemaManifest"] @("ListObject", $tableName, $table.Parent.Name, ($columnNames -join "|"), "2.0.0"))
    }
    foreach ($sheet in $workbook.Worksheets) {
        [void](Add-TableRow $tables["tblSchemaManifest"] @("Worksheet", $sheet.CodeName, $sheet.Name, "", "2.0.0"))
    }

    $welcome.Cells.Clear()
    $welcome.Range("B2:H2").Merge()
    $welcome.Range("B2").Value2 = "Third-Party Assessment Tracker"
    $welcome.Range("B2").Font.Size = 24
    $welcome.Range("B2").Font.Bold = $true
    $welcome.Range("B4:H6").Merge()
    $welcome.Range("B4").Value2 = "If you can read this but buttons do not work, macros are disabled. Close the file, confirm the publisher/signature or approved Trusted Location with IT, then reopen. New Outlook is not supported."
    $welcome.Range("B4").WrapText = $true
    $welcome.Range("B8").Value2 = "Windows identity: checked when macros run"
    $welcome.Range("B9").Value2 = "Access: fail-closed until startup checks pass"
    $welcome.Range("B10:H10").Merge()
    $welcome.Range("B:H").ColumnWidth = 18
    $welcome.Rows.Item(4).RowHeight = 72
    $welcome.Activate() | Out-Null
    $excel.ActiveWindow.DisplayGridlines = $false

    $dashboard.Range("A1:H1").Merge()
    $dashboard.Range("A1").Value2 = "Third-Party Assessment Dashboard"
    $dashboard.Range("A1").Font.Size = 20
    $dashboard.Range("A1").Font.Bold = $true
    $dashboardLabels = @("Active assessments","","Completed","","Due","","Overdue","")
    for ($i = 0; $i -lt $dashboardLabels.Count; $i++) { $dashboard.Cells.Item(3, $i + 1).Value2 = $dashboardLabels[$i] }
    $dashboard.Range("A4:H5").Interior.Color = 15921906
    $dashboard.Range("B4,D4,F4,H4").Font.Size = 20
    $dueHeaders = @("Assessment ID","Vendor","Cycle","Owner","Owner Email","Status","Next Due","Days Overdue")
    for ($i = 0; $i -lt $dueHeaders.Count; $i++) { $dashboard.Cells.Item(12, $i + 1).Value2 = $dueHeaders[$i] }
    $dashboard.Range("A12:H12").Font.Bold = $true
    $dashboard.Range("A:H").ColumnWidth = 18
    $dashboard.Range("B:B").ColumnWidth = 28
    $dashboard.Range("E:E").ColumnWidth = 30
    $dashboard.Range("J:K").EntireColumn.Hidden = $true
    Add-Button $dashboard "mut_AddVendor" "Add vendor" "modDashboard.OpenVendorForm" 20 155
    Add-Button $dashboard "mut_AddAssessment" "Add assessment" "modDashboard.OpenAssessmentForm" 145 155
    Add-Button $dashboard "mut_Import" "Import" "modImport.StartImport" 270 155
    Add-Button $dashboard "mut_Reconcile" "Reconcile sent" "modEmail.Reconcile" 395 155
    Add-Button $dashboard "mut_BatchDue" "Prepare all due" "modFollowup.PrepareAllDue" 520 155
    Add-Button $dashboard "mgr_Users" "Users" "modUsers.ShowUsersForm" 645 155
    Add-Button $dashboard "mgr_Settings" "Settings" "modDashboard.OpenSettingsForm" 770 155

    $vendorsView.Range("A1:H1").Merge()
    $vendorsView.Range("A1").Value2 = "Active Vendors"
    $vendorsView.Range("A1").Font.Size = 18
    $vendorsView.Range("A2").Value2 = "Double-click a row to edit. Business-owner changes affect defaults only, not existing assessment snapshots."
    $vendorsView.Range("A2:H2").Merge()
    $vendorsView.Range("A:H").ColumnWidth = 18
    $vendorsView.Range("B:B").ColumnWidth = 28
    $vendorsView.Range("H:H").ColumnWidth = 36

    $assessmentsView.Range("A1:N1").Merge()
    $assessmentsView.Range("A1").Value2 = "Active Assessments"
    $assessmentsView.Range("A1").Font.Size = 18
    $assessmentsView.Range("A2").Value2 = "Double-click a row to edit cycle/owner snapshot/details. Use Status for all status/date changes."
    $assessmentsView.Range("A2:N2").Merge()
    $assessmentsView.Range("A:N").ColumnWidth = 16
    $assessmentsView.Range("C:C").ColumnWidth = 26
    $assessmentsView.Range("F:F").ColumnWidth = 28
    Add-Button $assessmentsView "mut_Status" "Set status" "modDashboard.OpenStatusForm" 20 55
    Add-Button $assessmentsView "mut_Initial" "Prepare initial" "modDashboard.PrepareSelectedInitial" 145 55
    Add-Button $assessmentsView "mut_Followup" "Prepare follow-up" "modDashboard.PrepareSelectedFollowup" 270 55
    Add-Button $assessmentsView "mut_Meeting" "Log meeting" "modDashboard.LogSelectedMeeting" 395 55

    $dashboard.Activate() | Out-Null
    $excel.ActiveWindow.SplitRow = 12
    $excel.ActiveWindow.FreezePanes = $true
    $excel.ActiveWindow.DisplayGridlines = $false
    $vendorsView.Activate() | Out-Null
    $excel.ActiveWindow.SplitRow = 4
    $excel.ActiveWindow.FreezePanes = $true
    $excel.ActiveWindow.DisplayGridlines = $false
    $assessmentsView.Activate() | Out-Null
    $excel.ActiveWindow.SplitRow = 4
    $excel.ActiveWindow.FreezePanes = $true
    $excel.ActiveWindow.DisplayGridlines = $false
    $welcome.Activate() | Out-Null

    $importStaging.Visible = 0
    foreach ($sheet in @($dataVendors,$dataAssessments,$dataEmailEvents,$dataUsers,$dataAudit,$settings,$schema)) { $sheet.Visible = 2 }

    Add-ReferenceIfMissing $project "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}" 2 8 "Microsoft Office Object Library"
    Add-ReferenceIfMissing $project "{00020430-0000-0000-C000-000000000046}" 2 0 "OLE Automation"

    foreach ($module in $requiredModules) {
        [void]$project.VBComponents.Import((Join-Path $sourceDirectory ($module + ".bas")))
    }
    foreach ($form in $requiredForms) {
        [void]$project.VBComponents.Import((Join-Path $sourceDirectory ($form + ".frm")))
    }
    foreach ($entry in $documentSources.GetEnumerator()) {
        Install-DocumentCode $project $entry.Key (Join-Path $sourceDirectory $entry.Value)
    }

    $excel.EnableEvents = $true
    $workbook.Save()
    Write-Host "Running first structural-only self-test gate..."
    $firstMacroBookName = $workbook.Name.Replace("'", "''")
    $firstPass = [bool]$excel.Run(("'{0}'!modEnv.SelfTest" -f $firstMacroBookName), $false, $true)
    if (-not $firstPass) { throw "First structural-only modEnv.SelfTest gate failed. Inspect the workbook schema and data integrity." }
    $workbook.Save()
    $workbook.Close($true)
    Release-ComObject $workbook
    $workbook = $null

    Write-Host "Reopening saved workbook and running final structural-only self-test gate..."
    $excel.EnableEvents = $false
    $reopened = $excel.Workbooks.Open($OutputPath, 0, $false)
    $finalMacroBookName = $reopened.Name.Replace("'", "''")
    $finalPass = [bool]$excel.Run(("'{0}'!modEnv.SelfTest" -f $finalMacroBookName), $false, $true)
    if (-not $finalPass) { throw "Reopen structural-only modEnv.SelfTest gate failed. The build is not accepted." }
    $reopened.Save()
    $reopened.Close($true)
    Release-ComObject $reopened
    $reopened = $null
    $excel.EnableEvents = $true
    Write-Host "PASS: built, saved, reopened, and self-tested: $OutputPath" -ForegroundColor Green
    Write-Host "Initial Manager: $managerIdentity"
    Write-Host "Internal domain allowlist: $InternalDomain"
} catch {
    Write-Error ("BUILD FAILED: " + $_.Exception.Message)
    throw
} finally {
    if ($null -ne $reopened) { try { $reopened.Close($false) } catch {}; Release-ComObject $reopened }
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {}; Release-ComObject $workbook }
    if ($null -ne $excel) {
        try { if ($null -ne $oldAutomationSecurity) { $excel.AutomationSecurity = $oldAutomationSecurity } } catch {}
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
