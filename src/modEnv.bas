Attribute VB_Name = "modEnv"
Option Explicit

Public Function Check(Optional ByVal showReport As Boolean = True, Optional ByVal structuralOnly As Boolean = False) As Boolean
    Dim report As String
    Dim valid As Boolean
    If Not structuralOnly Then modIdentity.ClearEnvironmentBlock
    valid = RunChecks(report, structuralOnly)
    If Not structuralOnly And Not valid Then modIdentity.BlockMutations "Environment self-test failed:" & vbCrLf & report
    Check = valid
    If showReport Then
        If structuralOnly Then
            If valid Then MsgBox "Structural check PASS." & vbCrLf & SupportInfo(), vbInformation, modUtil.APP_NAME Else MsgBox "Structural check FAIL:" & vbCrLf & report, vbCritical, modUtil.APP_NAME
        Else
            If valid Then MsgBox "Environment check PASS." & vbCrLf & SupportInfo(), vbInformation, modUtil.APP_NAME Else MsgBox "Environment check FAIL. Mutations are blocked:" & vbCrLf & report, vbCritical, modUtil.APP_NAME
        End If
    End If
    If Not structuralOnly Then modDashboard.ApplyRoleBasedUI
End Function

Public Function SelfTest(Optional ByVal showReport As Boolean = True, Optional ByVal structuralOnly As Boolean = False) As Boolean
    SelfTest = Check(showReport, structuralOnly)
End Function

Private Function RunChecks(ByRef report As String, ByVal structuralOnly As Boolean) As Boolean
    Dim ok As Boolean
    Dim failureReason As String
    ok = True
    If modSettings.GetSetting("SchemaVersion", vbNullString) <> modUtil.SCHEMA_VERSION Then AddFailure report, ok, "SchemaVersion mismatch; migration is required."
    If Not CheckRequiredSettings(report) Then ok = False
    If Not CheckTables(report) Then ok = False
    If Not CheckWorkbookStructure(report) Then ok = False
    If Not CheckUsers(report) Then ok = False
    If Not CheckDataIntegrity(report) Then ok = False
    If Not structuralOnly Then
        If Not CheckPath() Then AddFailure report, ok, "Workbook path resembles a browser temp/quarantine location or is unsaved."
        If ThisWorkbook.ReadOnly Then
            AddFailure report, ok, "Workbook is read-only (Browse mode)."
        ElseIf Not modBackup.TestBackupFolder(failureReason) Then
            AddFailure report, ok, "Workbook/Backups folder is not writable: " & failureReason
        End If
        If Not CheckClassicOutlook(failureReason) Then AddFailure report, ok, "Classic Outlook/MAPI unavailable: " & failureReason
    End If
    RunChecks = ok
End Function

Private Function CheckWorkbookStructure(ByRef report As String) As Boolean
    Dim expectedSheets As Object
    Dim requiredTables As Variant
    Dim key As Variant
    Dim tableName As Variant
    Dim ws As Worksheet
    Dim manifest As ListObject
    Dim manifestRow As ListRow
    Dim ok As Boolean
    ok = True
    Set expectedSheets = CreateObject("Scripting.Dictionary")
    expectedSheets.Add "WELCOME", "shtWelcome"
    expectedSheets.Add "Dashboard", "shtDashboard"
    expectedSheets.Add "Vendors", "shtVendors"
    expectedSheets.Add "Assessments", "shtAssessments"
    expectedSheets.Add "ImportStaging", "shtImportStaging"
    expectedSheets.Add "Data_Vendors", "shtDataVendors"
    expectedSheets.Add "Data_Assessments", "shtDataAssessments"
    expectedSheets.Add "Data_EmailEvents", "shtDataEmailEvents"
    expectedSheets.Add "Data_Users", "shtDataUsers"
    expectedSheets.Add "Data_Audit", "shtDataAudit"
    expectedSheets.Add "Settings", "shtSettings"
    expectedSheets.Add "Schema", "shtSchema"
    On Error GoTo Failed
    Set manifest = modUtil.GetTable(modUtil.TBL_SCHEMA)
    For Each key In expectedSheets.Keys
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(key))
        On Error GoTo Failed
        If ws Is Nothing Then
            AddFailure report, ok, "Missing worksheet " & CStr(key)
        ElseIf ws.CodeName <> CStr(expectedSheets(key)) Then
            AddFailure report, ok, "Worksheet " & CStr(key) & " has CodeName " & ws.CodeName & ", expected " & CStr(expectedSheets(key))
        End If
    Next key
    requiredTables = Array(modUtil.TBL_VENDORS, modUtil.TBL_ASSESSMENTS, modUtil.TBL_EMAIL_EVENTS, modUtil.TBL_USERS, modUtil.TBL_AUDIT, modUtil.TBL_SETTINGS, modUtil.TBL_IMPORT, modUtil.TBL_SCHEMA)
    For Each tableName In requiredTables
        Set manifestRow = modUtil.FindTextRow(manifest, "ObjectName", CStr(tableName))
        If manifestRow Is Nothing Then AddFailure report, ok, "Schema manifest is missing " & CStr(tableName)
    Next tableName
    CheckWorkbookStructure = ok
    Exit Function
Failed:
    AddFailure report, ok, "Workbook structure validation failed: " & Err.Description
    CheckWorkbookStructure = False
End Function

Private Function CheckRequiredSettings(ByRef report As String) As Boolean
    Dim required As Variant
    Dim item As Variant
    Dim lo As ListObject
    Dim lr As ListRow
    Dim ok As Boolean
    Dim namedItem As Name
    Dim namedRange As Range
    ok = True
    required = Array("FollowupDays", "OrgName", "EmailInitialSubject", "EmailInitialTemplate", "EmailFollowupSubject", "EmailFollowupTemplate", "SendingAccountSMTP", "EmailMode", "BatchCap", "InternalDomainAllowlist", "NextVendorID", "NextAssessmentID", "NextEventID", "SchemaVersion", "StaleBackupKeep", "ReconcileGraceHours", "BootstrapState", "LastBackupDate")
    On Error GoTo Failed
    Set lo = modUtil.GetTable(modUtil.TBL_SETTINGS)
    For Each item In required
        Set lr = modUtil.FindTextRow(lo, "SettingName", CStr(item))
        If lr Is Nothing Then
            AddFailure report, ok, "Missing required setting " & CStr(item)
        Else
            Set namedItem = Nothing
            Set namedRange = Nothing
            On Error Resume Next
            Set namedItem = ThisWorkbook.Names("cfg" & CStr(item))
            If Not namedItem Is Nothing Then Set namedRange = namedItem.RefersToRange
            On Error GoTo Failed
            If namedRange Is Nothing Then
                AddFailure report, ok, "Missing/broken named range cfg" & CStr(item)
            ElseIf namedRange.Parent.Name <> "Settings" Or namedRange.Column <> 2 Or namedRange.Row <> lr.Range.Row Then
                AddFailure report, ok, "Named range cfg" & CStr(item) & " does not point to its SettingValue cell."
            End If
        End If
    Next item
    CheckRequiredSettings = ok
    Exit Function
Failed:
    AddFailure report, ok, "Settings validation failed: " & Err.Description
    CheckRequiredSettings = False
End Function

Private Function CheckTables(ByRef report As String) As Boolean
    Dim required As Object
    Dim key As Variant
    Dim columns As Variant
    Dim lo As ListObject
    Dim i As Long
    Dim ok As Boolean
    Set required = CreateObject("Scripting.Dictionary")
    required.Add modUtil.TBL_VENDORS, Split("VendorID|VendorName|BusinessOwnerName|BusinessOwnerEmail|VendorContactPerson|VendorContactEmail|VendorContactPhone|Notes|CreatedBy|CreatedOn|ModifiedBy|ModifiedOn|Archived|ArchivedOn|ArchivedBy", "|")
    required.Add modUtil.TBL_ASSESSMENTS, Split("AssessmentID|VendorID|Cycle|BusinessOwnerName|BusinessOwnerEmail|Status|SubmissionDate|CompletedDate|MeetingsConducted|Notes|CreatedBy|CreatedOn|ModifiedBy|ModifiedOn|Archived|ArchivedOn|ArchivedBy", "|")
    required.Add modUtil.TBL_EMAIL_EVENTS, Split("EventID|AssessmentID|Kind|State|CorrelationToken|Recipient|Subject|SendingAccount|DraftEntryID|StoreID|InternetMessageID|PreparedOn|SentOn|PreparedBy|LastStateOn|ErrorDetail", "|")
    required.Add modUtil.TBL_USERS, Split("WindowsUsername|DisplayName|Role|Active|CreatedBy|CreatedOn", "|")
    required.Add modUtil.TBL_AUDIT, Split("Timestamp|WindowsUsername|Action|EntityType|EntityID|Detail", "|")
    required.Add modUtil.TBL_SETTINGS, Split("SettingName|SettingValue|Description", "|")
    required.Add modUtil.TBL_IMPORT, Split("SourceRow|ImportMode|Action|Valid|Errors|VendorName|BusinessOwnerName|BusinessOwnerEmail|VendorContactPerson|VendorContactEmail|VendorContactPhone|Notes|Cycle", "|")
    required.Add modUtil.TBL_SCHEMA, Split("ObjectType|ObjectName|SheetName|Columns|SchemaVersion", "|")
    ok = True
    For Each key In required.Keys
        On Error Resume Next
        Set lo = Nothing
        Set lo = modUtil.GetTable(CStr(key))
        If Err.Number <> 0 Or lo Is Nothing Then
            AddFailure report, ok, "Missing table " & CStr(key)
            Err.Clear
        Else
            columns = required(key)
            For i = LBound(columns) To UBound(columns)
                If Not HasColumn(lo, CStr(columns(i))) Then AddFailure report, ok, lo.Name & " missing column " & CStr(columns(i))
            Next i
        End If
        On Error GoTo 0
    Next key
    CheckTables = ok
End Function

Private Function HasColumn(ByVal lo As ListObject, ByVal columnName As String) As Boolean
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = lo.ListColumns(columnName)
    HasColumn = Not lc Is Nothing
    On Error GoTo 0
End Function

Private Function CheckUsers(ByRef report As String) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    Dim seen As Object
    Dim identityValue As String
    Dim roleValue As String
    Dim managerCount As Long
    Dim ok As Boolean
    ok = True
    If modSettings.GetSetting("BootstrapState", vbNullString) <> "Bootstrapped" Then AddFailure report, ok, "BootstrapState is not exactly Bootstrapped."
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    Set lo = modUtil.GetTable(modUtil.TBL_USERS)
    For Each lr In lo.ListRows
        identityValue = Trim$(CStr(modUtil.RowValue(lr, lo, "WindowsUsername")))
        roleValue = Trim$(CStr(modUtil.RowValue(lr, lo, "Role")))
        If Len(identityValue) = 0 Or InStr(identityValue, "\") <= 1 Then AddFailure report, ok, "Invalid WindowsUsername in tblUsers."
        If Not IsRecognizedBoolean(modUtil.RowValue(lr, lo, "Active")) Then AddFailure report, ok, "Invalid Active flag for " & identityValue
        If seen.Exists(identityValue) Then AddFailure report, ok, "Duplicate WindowsUsername: " & identityValue Else seen(identityValue) = True
        If roleValue <> "Assessor" And roleValue <> "Manager" Then AddFailure report, ok, "Invalid user role for " & identityValue
        If modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Active"), False) And roleValue = "Manager" Then managerCount = managerCount + 1
    Next lr
    If lo.ListRows.count = 0 Or managerCount = 0 Then AddFailure report, ok, "Bootstrapped workbook requires at least one active Manager."
    CheckUsers = ok
End Function

Private Function IsRecognizedBoolean(ByVal value As Variant) As Boolean
    Dim s As String
    If VarType(value) = vbBoolean Then IsRecognizedBoolean = True: Exit Function
    s = LCase$(Trim$(CStr(value)))
    IsRecognizedBoolean = (s = "true" Or s = "false" Or s = "yes" Or s = "no" Or s = "1" Or s = "0")
End Function

Private Function CheckDataIntegrity(ByRef report As String) As Boolean
    Dim vendors As ListObject, assessments As ListObject, events As ListObject
    Dim lr As ListRow, parent As ListRow
    Dim pairs As Object, ids As Object
    Dim key As String, reason As String, stateValue As String, kindValue As String
    Dim ok As Boolean
    Dim guidTokens As Object, activeEvents As Object, initialSent As Object
    On Error GoTo Corrupt
    ok = True
    Set vendors = modUtil.GetTable(modUtil.TBL_VENDORS)
    Set assessments = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    Set events = modUtil.GetTable(modUtil.TBL_EMAIL_EVENTS)
    Set ids = CreateObject("Scripting.Dictionary")
    For Each lr In vendors.ListRows
        key = CStr(modUtil.RowValue(lr, vendors, "VendorID"))
        If Not IsNumeric(key) Or Val(key) < 1 Then AddFailure report, ok, "Invalid VendorID " & key
        If Len(Trim$(CStr(modUtil.RowValue(lr, vendors, "VendorName")))) = 0 Then AddFailure report, ok, "Vendor " & key & " has no name."
        If Len(Trim$(CStr(modUtil.RowValue(lr, vendors, "BusinessOwnerName")))) = 0 Or Not modUtil.IsEmailAddress(CStr(modUtil.RowValue(lr, vendors, "BusinessOwnerEmail"))) Then AddFailure report, ok, "Vendor " & key & " has an invalid default owner."
        If Not IsRecognizedBoolean(modUtil.RowValue(lr, vendors, "Archived")) Then AddFailure report, ok, "Vendor " & key & " has an invalid Archived flag."
        If ids.Exists(key) Then AddFailure report, ok, "Duplicate VendorID " & key Else ids(key) = True
    Next lr
    Set pairs = CreateObject("Scripting.Dictionary"): pairs.CompareMode = 1
    Set ids = CreateObject("Scripting.Dictionary")
    For Each lr In assessments.ListRows
        key = CStr(modUtil.RowValue(lr, assessments, "AssessmentID"))
        If Not IsNumeric(key) Or Val(key) < 1 Then AddFailure report, ok, "Invalid AssessmentID " & key
        If ids.Exists(key) Then AddFailure report, ok, "Duplicate AssessmentID " & key Else ids(key) = True
        Set parent = modUtil.FindLongRow(vendors, "VendorID", CLng(modUtil.RowValue(lr, assessments, "VendorID")))
        If parent Is Nothing Then AddFailure report, ok, "Orphan assessment " & key
        If Len(Trim$(CStr(modUtil.RowValue(lr, assessments, "BusinessOwnerName")))) = 0 Or Not modUtil.IsEmailAddress(CStr(modUtil.RowValue(lr, assessments, "BusinessOwnerEmail"))) Then AddFailure report, ok, "Invalid owner snapshot on assessment " & key
        If Len(Trim$(CStr(modUtil.RowValue(lr, assessments, "Cycle")))) = 0 Then AddFailure report, ok, "Assessment " & key & " has a blank Cycle."
        If Not IsRecognizedBoolean(modUtil.RowValue(lr, assessments, "Archived")) Then AddFailure report, ok, "Assessment " & key & " has an invalid Archived flag."
        If Not modStatus.ValidateInvariant(CStr(modUtil.RowValue(lr, assessments, "Status")), modUtil.RowValue(lr, assessments, "SubmissionDate"), modUtil.RowValue(lr, assessments, "CompletedDate"), reason) Then AddFailure report, ok, "Assessment " & key & ": " & reason
        key = CStr(modUtil.RowValue(lr, assessments, "VendorID")) & "|" & Trim$(CStr(modUtil.RowValue(lr, assessments, "Cycle")))
        If pairs.Exists(key) Then AddFailure report, ok, "Duplicate (VendorID,Cycle): " & key Else pairs(key) = True
    Next lr
    Set ids = CreateObject("Scripting.Dictionary")
    Set guidTokens = CreateObject("Scripting.Dictionary"): guidTokens.CompareMode = 1
    Set activeEvents = CreateObject("Scripting.Dictionary"): activeEvents.CompareMode = 1
    Set initialSent = CreateObject("Scripting.Dictionary")
    For Each lr In events.ListRows
        If CStr(modUtil.RowValue(lr, events, "Kind")) = "Initial" And CStr(modUtil.RowValue(lr, events, "State")) = "Sent" And IsDate(modUtil.RowValue(lr, events, "SentOn")) Then initialSent(CStr(modUtil.RowValue(lr, events, "AssessmentID"))) = True
    Next lr
    For Each lr In events.ListRows
        key = CStr(modUtil.RowValue(lr, events, "EventID"))
        If Not IsNumeric(key) Or Val(key) < 1 Then AddFailure report, ok, "Invalid EventID " & key
        If ids.Exists(key) Then AddFailure report, ok, "Duplicate EventID " & key Else ids(key) = True
        Set parent = modUtil.FindLongRow(assessments, "AssessmentID", CLng(modUtil.RowValue(lr, events, "AssessmentID")))
        If parent Is Nothing Then AddFailure report, ok, "Orphan email event " & key
        If Not IsGuidText(CStr(modUtil.RowValue(lr, events, "CorrelationToken"))) Then
            AddFailure report, ok, "Email event missing GUID token: " & key
        ElseIf guidTokens.Exists(CStr(modUtil.RowValue(lr, events, "CorrelationToken"))) Then
            AddFailure report, ok, "Duplicate email GUID token on event " & key
        Else
            guidTokens(CStr(modUtil.RowValue(lr, events, "CorrelationToken"))) = True
        End If
        stateValue = CStr(modUtil.RowValue(lr, events, "State"))
        If InStr(1, "|Prepared|DraftCreated|Queued|Sent|Unresolved|Failed|Cancelled|", "|" & stateValue & "|", vbBinaryCompare) = 0 Then AddFailure report, ok, "Invalid email state on event " & key
        kindValue = CStr(modUtil.RowValue(lr, events, "Kind"))
        If kindValue <> "Initial" And kindValue <> "Followup" Then AddFailure report, ok, "Invalid email kind on event " & key
        If Not modUtil.IsEmailAddress(CStr(modUtil.RowValue(lr, events, "Recipient"))) Then AddFailure report, ok, "Invalid recipient on event " & key
        If stateValue = "Sent" And Not IsDate(modUtil.RowValue(lr, events, "SentOn")) Then AddFailure report, ok, "Sent event has no SentOn: " & key
        If kindValue = "Followup" And stateValue = "Sent" And Not initialSent.Exists(CStr(modUtil.RowValue(lr, events, "AssessmentID"))) Then AddFailure report, ok, "Sent follow-up has no confirmed Initial event: " & key
        If stateValue = "Prepared" Or stateValue = "DraftCreated" Or stateValue = "Queued" Or stateValue = "Unresolved" Then
            key = CStr(modUtil.RowValue(lr, events, "AssessmentID")) & "|" & kindValue
            If activeEvents.Exists(key) Then AddFailure report, ok, "Duplicate pending/unresolved email event: " & key Else activeEvents(key) = True
        End If
    Next lr
    CheckDataIntegrity = ok
    Exit Function
Corrupt:
    AddFailure report, ok, "Data integrity scan could not parse a row: " & Err.Description
    CheckDataIntegrity = False
End Function

Private Function IsGuidText(ByVal value As String) As Boolean
    value = Trim$(value)
    IsGuidText = (Len(value) = 38 And Left$(value, 1) = "{" And Right$(value, 1) = "}" And Mid$(value, 10, 1) = "-" And Mid$(value, 15, 1) = "-" And Mid$(value, 20, 1) = "-" And Mid$(value, 25, 1) = "-")
End Function

Private Function CheckClassicOutlook(ByRef failureReason As String) As Boolean
    Dim outlookApp As Object
    Dim ns As Object
    On Error GoTo Failed
    Set outlookApp = CreateObject("Outlook.Application")
    Set ns = outlookApp.GetNamespace("MAPI")
    If ns.Stores.count < 1 Or outlookApp.Session.Accounts.count < 1 Then Err.Raise vbObjectError + 2000, , "No configured MAPI store/account."
    CheckClassicOutlook = True
    Exit Function
Failed:
    failureReason = Err.Description
End Function

Private Function CheckPath() As Boolean
    Dim p As String
    p = LCase$(ThisWorkbook.FullName)
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    If InStr(p, "content.outlook") > 0 Or InStr(p, "temporary internet files") > 0 Or InStr(p, "\appdata\local\temp\") > 0 Or InStr(p, "\downloads\") > 0 Then Exit Function
    CheckPath = True
End Function

Private Sub AddFailure(ByRef report As String, ByRef ok As Boolean, ByVal messageText As String)
    ok = False
    If Len(report) > 0 Then report = report & vbCrLf
    report = report & "- " & messageText
End Sub

Private Function SupportInfo() As String
#If Win64 Then
    SupportInfo = "Office bitness: 64-bit"
#Else
    SupportInfo = "Office bitness: 32-bit"
#End If
End Function
