Option Explicit

' Pure-VBA, offline bootstrap for ThirdParty_Assessment_Tracker.xlsm.
' Paste this standard module into a saved macro-enabled workbook beside .\src,
' enable "Trust access to the VBA project object model", and run BuildTracker.

Private Const TARGET_FILE As String = "ThirdParty_Assessment_Tracker.xlsm"
Private Const LOG_SHEET As String = "BuildLog"
Private Const SCHEMA_VERSION As String = "2.0.0"
Private Const XL_OPEN_XML_WORKBOOK_MACRO_ENABLED As Long = 52
Private Const XL_SRC_RANGE As Long = 1
Private Const XL_YES As Long = 1
Private Const MSO_SHAPE_ROUNDED_RECTANGLE As Long = 5
Private Const MSO_FALSE As Long = 0
Private Const XL_SHEET_HIDDEN As Long = 0
Private Const XL_SHEET_VERY_HIDDEN As Long = 2

Private mLogSheet As Worksheet
Private mLogRow As Long
Private mLogFile As Integer
Private mLogFileOpen As Boolean
Private mLogPath As String
Private mCurrentStep As String
Private mTarget As Workbook
Private mTargetSaved As Boolean

Public Sub BuildTracker()
    Dim sourceDirectory As String
    Dim outputPath As String
    Dim project As Object
    Dim managerIdentity As String
    Dim managerDisplayName As String
    Dim internalDomain As String
    Dim importsOK As Boolean
    Dim reportShimOK As Boolean
    Dim selfTestOK As Boolean
    Dim finalOK As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo FatalError
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    SetStep "Initialize BuildLog worksheet and text file"
    InitializeLogging
    LogEnvironment
    LogMessage "INFO", "Bootstrap workbook: " & ThisWorkbook.FullName

    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 4100, "BuildTracker", _
            "The bootstrap workbook must be saved before running BuildTracker. Save it as a macro-enabled workbook next to the src folder."
    End If

    SetStep "Resolve source directory"
    sourceDirectory = ResolveSourceDirectory()
    LogMessage "INFO", "Resolved src path: " & sourceDirectory

    SetStep "Validate all required source files"
    If Not ValidateRequiredFiles(sourceDirectory) Then
        Err.Raise vbObjectError + 4101, "BuildTracker", _
            "Source validation failed. One or more required files are missing; see the BuildLog."
    End If

    SetStep "Resolve output path"
    outputPath = ThisWorkbook.Path & Application.PathSeparator & TARGET_FILE
    LogMessage "INFO", "Target path: " & outputPath
    If StrComp(outputPath, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 4102, "BuildTracker", _
            "The bootstrap workbook itself is named " & TARGET_FILE & ". Rename the bootstrap workbook so the generated target has a different file."
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    SetStep "Create target workbook"
    Set mTarget = Workbooks.Add
    LogMessage "PASS", "Created new workbook: " & mTarget.Name

    SetStep "Check VBOM trust on target workbook"
    If Not GetTrustedProject(mTarget, project) Then GoTo FailedWithoutRaise

    SetStep "Collect bootstrap identity and domain settings"
    GetBootstrapInputs managerIdentity, managerDisplayName, internalDomain

    SetStep "Create workbook sheets and assign CodeNames"
    BuildWorksheets mTarget, project

    SetStep "Create tables, settings, named ranges, schema manifest, and bootstrap rows"
    BuildDataModel mTarget, managerIdentity, managerDisplayName, internalDomain

    SetStep "Format user-facing worksheets and add buttons"
    FormatWorkbook mTarget

    SetStep "Set worksheet visibility"
    SetWorksheetVisibility mTarget

    SetStep "Add optional VBA references"
    AddReferenceIfMissing project, "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 8, "Microsoft Office Object Library"
    AddReferenceIfMissing project, "{00020430-0000-0000-C000-000000000046}", 2, 0, "OLE Automation"

    SetStep "Import all VBA source files"
    importsOK = ImportAllCode(project, sourceDirectory)

    SetStep "Install modEnv diagnostic report shim"
    reportShimOK = InstallSelfTestReportShim(project)
    If Not importsOK Then
        LogMessage "FAIL", "One or more code imports failed. The remaining import categories were still attempted for a complete diagnostic log."
    End If

    SetStep "Prepare target output path"
    PrepareOutputPath outputPath

    SetStep "Save generated macro-enabled workbook"
    mTarget.SaveAs Filename:=outputPath, FileFormat:=XL_OPEN_XML_WORKBOOK_MACRO_ENABLED
    mTargetSaved = True
    LogMessage "PASS", "Saved target workbook: " & outputPath

    If reportShimOK Then
        SetStep "Run modEnv.SelfTest(False, True) structural gate"
        selfTestOK = RunStructuralSelfTest(mTarget)
    Else
        selfTestOK = False
        LogMessage "FAIL", "Structural self-test could not run because the required modEnv diagnostic-shim installation failed."
    End If

    SetStep "Save target after structural self-test"
    mTarget.Save
    LogMessage "PASS", "Saved target after structural self-test."

    finalOK = importsOK And reportShimOK And selfTestOK
    If Not selfTestOK Then
        LogMessage "FAIL", "SELF-TEST RETURNED FALSE. The workbook was preserved for diagnosis; this condition did not abort or delete the build."
    End If

    SetStep "Close generated workbook"
    mTarget.Close SaveChanges:=True
    Set mTarget = Nothing

    If finalOK Then
        LogBanner True, "Build completed and modEnv structural self-test passed."
    Else
        LogBanner False, "Build completed with failures. Review the detailed log above and copy the entire text file back for iteration."
    End If
    GoTo CleanExit

FailedWithoutRaise:
    finalOK = False
    LogBanner False, "Build aborted cleanly. Review the detailed error and instructions above."
    GoTo CleanExit

FatalError:
    errNumber = Err.Number
    errDescription = Err.Description
    SafeLogError mCurrentStep, errNumber, errDescription
    SafeLogBanner False, "BUILD FAILED at step: " & mCurrentStep
    finalOK = False

CleanExit:
    On Error Resume Next
    If Not mTarget Is Nothing Then mTarget.Close SaveChanges:=mTargetSaved
    Set mTarget = Nothing
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    PersistBootstrapWorkbook
    CloseLogFile
    On Error GoTo 0

    If finalOK Then
        MsgBox "PASS: the tracker was built and structurally self-tested." & vbCrLf & _
               ThisWorkbook.Path & Application.PathSeparator & TARGET_FILE & vbCrLf & _
               "Build log: " & mLogPath, vbInformation, "Tracker bootstrap"
    Else
        MsgBox "FAIL: the tracker bootstrap did not fully pass." & vbCrLf & _
               "Open the BuildLog sheet or copy this file back for support:" & vbCrLf & _
               mLogPath, vbCritical, "Tracker bootstrap"
    End If
End Sub

Private Sub InitializeLogging()
    Dim existing As Worksheet
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    Set existing = Nothing
    On Error Resume Next
    Set existing = ThisWorkbook.Worksheets(LOG_SHEET)
    On Error GoTo Failed

    If existing Is Nothing Then
        Set mLogSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        mLogSheet.Name = LOG_SHEET
    Else
        Set mLogSheet = existing
        mLogSheet.Cells.Clear
    End If

    mLogSheet.Cells(1, 1).Value2 = "Timestamp"
    mLogSheet.Cells(1, 2).Value2 = "Level"
    mLogSheet.Cells(1, 3).Value2 = "Message"
    mLogSheet.Range("A1:C1").Font.Bold = True
    mLogSheet.Columns("A").ColumnWidth = 22
    mLogSheet.Columns("B").ColumnWidth = 10
    mLogSheet.Columns("C").ColumnWidth = 110
    mLogSheet.Columns("C").WrapText = True
    mLogRow = 2

    If Len(ThisWorkbook.Path) > 0 Then
        mLogPath = ThisWorkbook.Path & Application.PathSeparator & _
                   "BuildLog_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
        mLogFile = FreeFile
        Open mLogPath For Output As #mLogFile
        mLogFileOpen = True
    Else
        mLogPath = "(not created: bootstrap workbook is unsaved)"
    End If

    LogMessage "INFO", String$(78, "=")
    LogMessage "INFO", "Third-Party Assessment Tracker - pure VBA bootstrap log"
    LogMessage "INFO", "Started: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    LogMessage "INFO", "Text log: " & mLogPath
    LogMessage "INFO", String$(78, "=")
    Exit Sub

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Raise errNumber, "InitializeLogging", errDescription
End Sub

Private Sub LogEnvironment()
    On Error GoTo Failed
    LogMessage "INFO", "Excel version: " & Application.Version
    LogMessage "INFO", "Excel build: " & Application.Build
    LogMessage "INFO", "Operating system: " & Application.OperatingSystem
#If Win64 Then
    LogMessage "INFO", "Office bitness: 64-bit (#If Win64=True)"
#Else
    LogMessage "INFO", "Office bitness: 32-bit (#If Win64=False)"
#End If
#If VBA7 Then
    LogMessage "INFO", "VBA host: VBA7"
#Else
    LogMessage "INFO", "VBA host: legacy VBA"
#End If
    Exit Sub
Failed:
    RaiseLoggedError "Log Excel version and bitness", Err.Number, Err.Description
End Sub

Private Sub LogMessage(ByVal levelText As String, ByVal messageText As String)
    Dim stamp As String
    stamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    If Not mLogSheet Is Nothing Then
        mLogSheet.Cells(mLogRow, 1).Value2 = stamp
        mLogSheet.Cells(mLogRow, 2).Value2 = levelText
        mLogSheet.Cells(mLogRow, 3).Value2 = messageText
        mLogRow = mLogRow + 1
    End If

    If mLogFileOpen Then
        Print #mLogFile, stamp & " [" & levelText & "] " & messageText
    End If
    DoEvents
End Sub

Private Sub LogBlock(ByVal levelText As String, ByVal heading As String, ByVal blockText As String)
    Dim normalized As String
    Dim lines As Variant
    Dim lineItem As Variant

    On Error GoTo Failed
    LogMessage levelText, heading
    normalized = Replace(blockText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    lines = Split(normalized, vbLf)
    For Each lineItem In lines
        LogMessage levelText, "    " & CStr(lineItem)
    Next lineItem
    Exit Sub
Failed:
    RaiseLoggedError "Write multi-line log block: " & heading, Err.Number, Err.Description
End Sub

Private Sub LogError(ByVal stepText As String, ByVal errNumber As Long, ByVal errDescription As String)
    LogMessage "ERROR", "Step='" & stepText & "' | Err.Number=" & CStr(errNumber) & _
               " | Err.Description=" & errDescription
End Sub

Private Sub LogBanner(ByVal passed As Boolean, ByVal detail As String)
    Dim stateText As String
    If passed Then stateText = "FINAL PASS" Else stateText = "FINAL FAIL"
    LogMessage IIf(passed, "PASS", "FAIL"), String$(78, "=")
    LogMessage IIf(passed, "PASS", "FAIL"), stateText & ": " & detail
    LogMessage IIf(passed, "PASS", "FAIL"), "Finished: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    LogMessage IIf(passed, "PASS", "FAIL"), String$(78, "=")
End Sub

Private Sub SafeLogError(ByVal stepText As String, ByVal errNumber As Long, ByVal errDescription As String)
    On Error Resume Next
    LogError stepText, errNumber, errDescription
    On Error GoTo 0
End Sub

Private Sub SafeLogBanner(ByVal passed As Boolean, ByVal detail As String)
    On Error Resume Next
    LogBanner passed, detail
    On Error GoTo 0
End Sub

Private Sub CloseLogFile()
    On Error Resume Next
    If mLogFileOpen Then Close #mLogFile
    mLogFileOpen = False
    On Error GoTo 0
End Sub

Private Sub PersistBootstrapWorkbook()
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Warning
    If Len(ThisWorkbook.Path) = 0 Then Exit Sub
    ThisWorkbook.Save
    Exit Sub
Warning:
    errNumber = Err.Number
    errDescription = Err.Description
    SafeLogError "Save bootstrap workbook with BuildLog worksheet", errNumber, errDescription
    Err.Clear
End Sub

Private Sub SetStep(ByVal stepText As String)
    mCurrentStep = stepText
    LogMessage "STEP", stepText
End Sub

Private Sub RaiseLoggedError(ByVal stepText As String, ByVal errNumber As Long, ByVal errDescription As String)
    LogError stepText, errNumber, errDescription
    Err.Raise errNumber, stepText, errDescription
End Sub

Private Function ResolveSourceDirectory() As String
    Dim preferred As String
    Dim entered As String
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    preferred = ThisWorkbook.Path & Application.PathSeparator & "src"
    If FolderExists(preferred) Then
        LogMessage "PASS", "Found preferred source folder beside bootstrap workbook."
        ResolveSourceDirectory = preferred
        Exit Function
    End If

    LogMessage "WARN", "Preferred source folder not found: " & preferred
    entered = InputBox( _
        "Enter the full path to the src folder containing the .bas, .cls, .frm, and placeholder .frx files.", _
        "Locate tracker src folder", preferred)
    entered = Trim$(entered)
    If Len(entered) = 0 Then
        Err.Raise vbObjectError + 4110, "ResolveSourceDirectory", "No source folder was supplied."
    End If
    If Left$(entered, 1) = Chr$(34) And Right$(entered, 1) = Chr$(34) Then
        entered = Mid$(entered, 2, Len(entered) - 2)
    End If
    Do While Len(entered) > 3 And Right$(entered, 1) = Application.PathSeparator
        entered = Left$(entered, Len(entered) - 1)
    Loop
    If Not FolderExists(entered) Then
        Err.Raise vbObjectError + 4111, "ResolveSourceDirectory", "The selected source folder does not exist: " & entered
    End If
    ResolveSourceDirectory = entered
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Resolve source directory", errNumber, errDescription
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    Dim attributes As Long
    On Error GoTo NotFound
    attributes = GetAttr(folderPath)
    FolderExists = ((attributes And vbDirectory) = vbDirectory)
    Exit Function
NotFound:
    FolderExists = False
    Err.Clear
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error GoTo NotFound
    FileExists = (Len(Dir$(filePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) > 0)
    Exit Function
NotFound:
    FileExists = False
    Err.Clear
End Function

Private Function RequiredModules() As Variant
    RequiredModules = Array( _
        "modIdentity", "modVendors", "modAssessments", "modStatus", "modEmail", "modFollowup", "modImport", _
        "modDashboard", "modUsers", "modSettings", "modEnv", "modBackup", "modAudit", "modUtil")
End Function

Private Function RequiredForms() As Variant
    RequiredForms = Array("frmVendor", "frmAssessment", "frmStatus", "frmImportReview", _
                          "frmUsers", "frmSettings", "frmBatchPreview")
End Function

Private Function DocumentComponents() As Variant
    DocumentComponents = Array( _
        Array("ThisWorkbook", "ThisWorkbook.cls"), _
        Array("shtWelcome", "shtWelcome.cls"), _
        Array("shtDashboard", "shtDashboard.cls"), _
        Array("shtVendors", "shtVendors.cls"), _
        Array("shtAssessments", "shtAssessments.cls"), _
        Array("shtImportStaging", "shtImportStaging.cls"), _
        Array("shtDataVendors", "shtDataVendors.cls"), _
        Array("shtDataAssessments", "shtDataAssessments.cls"), _
        Array("shtDataEmailEvents", "shtDataEmailEvents.cls"), _
        Array("shtDataUsers", "shtDataUsers.cls"), _
        Array("shtDataAudit", "shtDataAudit.cls"), _
        Array("shtSettings", "shtSettings.cls"), _
        Array("shtSchema", "shtSchema.cls"))
End Function

Private Function ValidateRequiredFiles(ByVal sourceDirectory As String) As Boolean
    Dim modules As Variant
    Dim forms As Variant
    Dim documents As Variant
    Dim item As Variant
    Dim pair As Variant
    Dim filePath As String
    Dim missingCount As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    modules = RequiredModules()
    forms = RequiredForms()
    documents = DocumentComponents()

    For Each item In modules
        filePath = sourceDirectory & Application.PathSeparator & CStr(item) & ".bas"
        LogPresence filePath, missingCount
    Next item

    For Each item In forms
        filePath = sourceDirectory & Application.PathSeparator & CStr(item) & ".frm"
        LogPresence filePath, missingCount
        filePath = sourceDirectory & Application.PathSeparator & CStr(item) & ".frx"
        LogPresence filePath, missingCount
    Next item

    For Each pair In documents
        filePath = sourceDirectory & Application.PathSeparator & CStr(pair(1))
        LogPresence filePath, missingCount
    Next pair

    If missingCount = 0 Then
        LogMessage "PASS", "Required-file validation passed. All 41 files are present (14 .bas, 7 .frm, 7 placeholder .frx, 13 .cls)."
        LogMessage "INFO", "The 7 placeholder .frx files were presence-validated only. They will not be copied or imported; each .frm is imported alone from an isolated TEMP folder."
        ValidateRequiredFiles = True
    Else
        LogMessage "FAIL", CStr(missingCount) & " required source file(s) are missing. No target workbook will be built."
        ValidateRequiredFiles = False
    End If
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Validate required source files", errNumber, errDescription
End Function

Private Sub LogPresence(ByVal filePath As String, ByRef missingCount As Long)
    On Error GoTo Failed
    If FileExists(filePath) Then
        LogMessage "FOUND", filePath
    Else
        missingCount = missingCount + 1
        LogMessage "MISSING", filePath
    End If
    Exit Sub
Failed:
    RaiseLoggedError "Check required file presence: " & filePath, Err.Number, Err.Description
End Sub

Private Sub PrepareOutputPath(ByVal outputPath As String)
    Dim response As VbMsgBoxResult
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    If Not FileExists(outputPath) Then Exit Sub

    LogMessage "WARN", "Output already exists: " & outputPath
    response = MsgBox( _
        "The generated tracker already exists:" & vbCrLf & outputPath & vbCrLf & vbCrLf & _
        "Replace that exact file? This cannot be undone by the bootstrap.", _
        vbExclamation Or vbYesNo Or vbDefaultButton2, "Replace existing tracker?")
    If response <> vbYes Then
        Err.Raise vbObjectError + 4120, "PrepareOutputPath", "Build cancelled because the output file already exists."
    End If
    Kill outputPath
    LogMessage "WARN", "User approved replacement; deleted existing output file: " & outputPath
    Exit Sub

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Prepare target output path", errNumber, errDescription
End Sub

Private Function GetTrustedProject(ByVal target As Workbook, ByRef project As Object) As Boolean
    Dim componentCount As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Denied
    Set project = target.VBProject
    componentCount = project.VBComponents.Count
    LogMessage "PASS", "VBOM trust status: ENABLED; target VBProject contains " & CStr(componentCount) & " initial component(s)."
    GetTrustedProject = True
    Exit Function

Denied:
    errNumber = Err.Number
    errDescription = Err.Description
    LogMessage "FAIL", "VBOM trust status: DISABLED or blocked by policy."
    LogBlock "FAIL", "Enable VBOM trust exactly as follows:", _
        "File > Options > Trust Center > Trust Center Settings > Macro Settings > " & _
        "'Trust access to the VBA project object model'" & vbCrLf & _
        "Close every Excel window, reopen the bootstrap workbook, and run BuildTracker again. " & _
        "If the option is disabled by Group Policy, contact corporate IT."
    LogError "Access target workbook VBProject.VBComponents", errNumber, errDescription
    GetTrustedProject = False
    Err.Clear
End Function

Private Sub GetBootstrapInputs(ByRef managerIdentity As String, _
                               ByRef managerDisplayName As String, _
                               ByRef internalDomain As String)
    Dim defaultIdentity As String
    Dim defaultDisplay As String
    Dim defaultDomain As String
    Dim slashPosition As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    defaultIdentity = Trim$(Environ$("USERDOMAIN")) & "\" & Trim$(Environ$("USERNAME"))
    managerIdentity = InputBox("Enter the initial Manager Windows identity as DOMAIN\Username.", _
                               "Initial Manager", defaultIdentity)
    If Len(Trim$(managerIdentity)) = 0 Then managerIdentity = defaultIdentity
    managerIdentity = UCase$(Trim$(managerIdentity))

    slashPosition = InStr(1, managerIdentity, "\", vbBinaryCompare)
    If slashPosition <= 1 Or slashPosition >= Len(managerIdentity) Or _
       InStr(slashPosition + 1, managerIdentity, "\", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 4130, "GetBootstrapInputs", _
            "Initial Manager must be exactly DOMAIN\USERNAME. Received: " & managerIdentity
    End If

    defaultDisplay = Mid$(managerIdentity, slashPosition + 1)
    managerDisplayName = Trim$(InputBox("Enter the initial Manager display name.", _
                                        "Initial Manager display name", defaultDisplay))
    If Len(managerDisplayName) = 0 Then managerDisplayName = defaultDisplay

    defaultDomain = Trim$(Environ$("USERDNSDOMAIN"))
    internalDomain = Trim$(InputBox( _
        "Enter the internal email domain used for the exact-recipient allowlist (for example, company.com).", _
        "Internal email domain", defaultDomain))
    If Len(internalDomain) = 0 Then internalDomain = defaultDomain
    If Len(internalDomain) = 0 Then
        internalDomain = "example.invalid"
        LogMessage "WARN", "USERDNSDOMAIN was unavailable and no domain was entered. InternalDomainAllowlist uses example.invalid and must be replaced by a Manager."
    End If
    Do While Left$(internalDomain, 1) = "@"
        internalDomain = Mid$(internalDomain, 2)
    Loop
    internalDomain = LCase$(Trim$(internalDomain))

    LogMessage "INFO", "Initial Manager identity: " & managerIdentity
    LogMessage "INFO", "Initial Manager display name: " & managerDisplayName
    LogMessage "INFO", "Internal domain allowlist: " & internalDomain
    Exit Sub

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Collect initial Manager and internal domain", errNumber, errDescription
End Sub

Private Function SheetDefinitions() As Variant
    SheetDefinitions = Array( _
        Array("WELCOME", "shtWelcome"), _
        Array("Dashboard", "shtDashboard"), _
        Array("Vendors", "shtVendors"), _
        Array("Assessments", "shtAssessments"), _
        Array("ImportStaging", "shtImportStaging"), _
        Array("Data_Vendors", "shtDataVendors"), _
        Array("Data_Assessments", "shtDataAssessments"), _
        Array("Data_EmailEvents", "shtDataEmailEvents"), _
        Array("Data_Users", "shtDataUsers"), _
        Array("Data_Audit", "shtDataAudit"), _
        Array("Settings", "shtSettings"), _
        Array("Schema", "shtSchema"))
End Function

Private Sub BuildWorksheets(ByVal target As Workbook, ByVal project As Object)
    Dim definitions As Variant
    Dim definition As Variant
    Dim ws As Worksheet
    Dim index As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    definitions = SheetDefinitions()

    Do While target.Worksheets.Count > 1
        target.Worksheets(target.Worksheets.Count).Delete
    Loop

    Set ws = target.Worksheets(1)
    ws.Name = CStr(definitions(0)(0))
    SetWorksheetCodeName project, ws, CStr(definitions(0)(1))
    LogMessage "PASS", "Created worksheet '" & ws.Name & "' with CodeName '" & ws.CodeName & "'."

    For index = 1 To UBound(definitions)
        Set ws = target.Worksheets.Add(After:=target.Worksheets(target.Worksheets.Count))
        ws.Name = CStr(definitions(index)(0))
        SetWorksheetCodeName project, ws, CStr(definitions(index)(1))
        LogMessage "PASS", "Created worksheet '" & ws.Name & "' with CodeName '" & ws.CodeName & "'."
    Next index
    Exit Sub

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Create worksheets and set CodeNames", errNumber, errDescription
End Sub

Private Sub SetWorksheetCodeName(ByVal project As Object, ByVal ws As Worksheet, ByVal codeName As String)
    Dim oldCodeName As String
    On Error GoTo Failed
    oldCodeName = ws.CodeName
    project.VBComponents.Item(oldCodeName).Name = codeName
    If ws.CodeName <> codeName Then
        Err.Raise vbObjectError + 4140, "SetWorksheetCodeName", _
            "Worksheet " & ws.Name & " retained CodeName " & ws.CodeName & "; expected " & codeName
    End If
    Exit Sub
Failed:
    RaiseLoggedError "Set CodeName '" & codeName & "' on worksheet '" & ws.Name & "'", Err.Number, Err.Description
End Sub

Private Function TableDefinitions() As Variant
    TableDefinitions = Array( _
        Array("tblVendors", "Data_Vendors", "VendorID|VendorName|BusinessOwnerName|BusinessOwnerEmail|VendorContactPerson|VendorContactEmail|VendorContactPhone|Notes|CreatedBy|CreatedOn|ModifiedBy|ModifiedOn|Archived|ArchivedOn|ArchivedBy"), _
        Array("tblAssessments", "Data_Assessments", "AssessmentID|VendorID|Cycle|BusinessOwnerName|BusinessOwnerEmail|Status|SubmissionDate|CompletedDate|MeetingsConducted|Notes|CreatedBy|CreatedOn|ModifiedBy|ModifiedOn|Archived|ArchivedOn|ArchivedBy"), _
        Array("tblEmailEvents", "Data_EmailEvents", "EventID|AssessmentID|Kind|State|CorrelationToken|Recipient|Subject|SendingAccount|DraftEntryID|StoreID|InternetMessageID|PreparedOn|SentOn|PreparedBy|LastStateOn|ErrorDetail"), _
        Array("tblUsers", "Data_Users", "WindowsUsername|DisplayName|Role|Active|CreatedBy|CreatedOn"), _
        Array("tblAudit", "Data_Audit", "Timestamp|WindowsUsername|Action|EntityType|EntityID|Detail"), _
        Array("tblSettings", "Settings", "SettingName|SettingValue|Description"), _
        Array("tblImportStaging", "ImportStaging", "SourceRow|ImportMode|Action|Valid|Errors|VendorName|BusinessOwnerName|BusinessOwnerEmail|VendorContactPerson|VendorContactEmail|VendorContactPhone|Notes|Cycle"), _
        Array("tblSchemaManifest", "Schema", "ObjectType|ObjectName|SheetName|Columns|SchemaVersion"))
End Function

Private Function SettingDefinitions(ByVal internalDomain As String) As Variant
    SettingDefinitions = Array( _
        Array("FollowupDays", "7", "Days after the last confirmed sent email before follow-up is due."), _
        Array("OrgName", "Your Organization", "Organization name used in templates."), _
        Array("EmailInitialSubject", "Third-Party Assessment: {VendorName} ({Cycle})", "Initial email subject template."), _
        Array("EmailInitialTemplate", "Hello {OwnerName}," & vbCrLf & vbCrLf & "Please begin the {Cycle} assessment for {VendorName} on behalf of {OrgName}." & vbCrLf & vbCrLf & "Thank you.", "Initial plain-text email body."), _
        Array("EmailFollowupSubject", "Follow-up: {VendorName} assessment ({Cycle})", "Follow-up email subject template."), _
        Array("EmailFollowupTemplate", "Hello {OwnerName}," & vbCrLf & vbCrLf & "This is a follow-up for the {Cycle} assessment of {VendorName} for {OrgName}." & vbCrLf & vbCrLf & "Thank you.", "Follow-up plain-text email body."), _
        Array("SendingAccountSMTP", "", "Exact Outlook account SMTP; blank uses the first configured account."), _
        Array("EmailMode", "Draft", "Draft by default. Direct is Manager-only and records Queued until reconciliation."), _
        Array("BatchCap", "50", "Maximum items in one batch."), _
        Array("InternalDomainAllowlist", internalDomain, "Semicolon-separated exact domains after @; no substring matching."), _
        Array("NextVendorID", "1", "Next stored monotonic vendor ID."), _
        Array("NextAssessmentID", "1", "Next stored monotonic assessment ID."), _
        Array("NextEventID", "1", "Next stored monotonic email event ID."), _
        Array("SchemaVersion", SCHEMA_VERSION, "Expected schema version."), _
        Array("StaleBackupKeep", "10", "Number of same-share recovery copies to retain."), _
        Array("ReconcileGraceHours", "24", "Hours before a missing Outlook item becomes Unresolved."), _
        Array("BootstrapState", "BootstrapPending", "Build-only state marker; runtime never grants Manager."), _
        Array("LastBackupDate", "", "Reserved support field; filenames enforce once-per-day backups."))
End Function

Private Sub BuildDataModel(ByVal target As Workbook, _
                           ByVal managerIdentity As String, _
                           ByVal managerDisplayName As String, _
                           ByVal internalDomain As String)
    Dim tableDefs As Variant
    Dim settingDefs As Variant
    Dim sheetDefs As Variant
    Dim definition As Variant
    Dim settingDef As Variant
    Dim sheetDef As Variant
    Dim tables As Object
    Dim settingRows As Object
    Dim lo As ListObject
    Dim lr As ListRow
    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim nowValue As Date
    Dim formulaText As String
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    Set tables = CreateObject("Scripting.Dictionary")
    Set settingRows = CreateObject("Scripting.Dictionary")
    tableDefs = TableDefinitions()

    For Each definition In tableDefs
        Set ws = target.Worksheets(CStr(definition(1)))
        Set lo = AddTable(ws, CStr(definition(0)), Split(CStr(definition(2)), "|"))
        tables.Add CStr(definition(0)), lo
        LogMessage "PASS", "Created table '" & lo.Name & "' on '" & ws.Name & "' with columns: " & CStr(definition(2))
    Next definition

    settingDefs = SettingDefinitions(internalDomain)
    For Each settingDef In settingDefs
        Set lr = AddTableRow(tables("tblSettings"), Array(settingDef(0), settingDef(1), settingDef(2)))
        rowNumber = lr.Range.Row
        settingRows.Add CStr(settingDef(0)), rowNumber
        LogMessage "PASS", "Created setting '" & CStr(settingDef(0)) & "' default='" & LogSafeValue(settingDef(1)) & "'."
    Next settingDef

    nowValue = Now
    Set lr = AddTableRow(tables("tblUsers"), _
        Array(managerIdentity, managerDisplayName, "Manager", True, "BUILD_SCRIPT", nowValue))
    LogMessage "PASS", "Provisioned initial active Manager in tblUsers: " & managerIdentity

    Set lr = AddTableRow(tables("tblAudit"), _
        Array(nowValue, managerIdentity, "BOOTSTRAP", "Workbook", "InitialManager", _
              "Build script provisioned initial Manager and changed BootstrapPending to Bootstrapped."))
    LogMessage "PASS", "Added BOOTSTRAP audit row for initial Manager."

    target.Worksheets("Settings").Cells(CLng(settingRows("BootstrapState")), 2).Value2 = "Bootstrapped"
    LogMessage "PASS", "Changed setting BootstrapState from BootstrapPending to Bootstrapped."

    For Each settingDef In settingDefs
        formulaText = "='Settings'!$B$" & CStr(settingRows(CStr(settingDef(0))))
        target.Names.Add Name:="cfg" & CStr(settingDef(0)), RefersTo:=formulaText
        LogMessage "PASS", "Created named range cfg" & CStr(settingDef(0)) & " -> " & formulaText
    Next settingDef

    For Each definition In tableDefs
        Set lo = tables(CStr(definition(0)))
        Set lr = AddTableRow(tables("tblSchemaManifest"), _
            Array("ListObject", lo.Name, lo.Parent.Name, JoinListColumns(lo), SCHEMA_VERSION))
        LogMessage "PASS", "Added schema-manifest ListObject row: " & lo.Name
    Next definition

    sheetDefs = SheetDefinitions()
    For Each sheetDef In sheetDefs
        Set ws = target.Worksheets(CStr(sheetDef(0)))
        Set lr = AddTableRow(tables("tblSchemaManifest"), _
            Array("Worksheet", ws.CodeName, ws.Name, "", SCHEMA_VERSION))
        LogMessage "PASS", "Added schema-manifest Worksheet row: " & ws.CodeName & " / " & ws.Name
    Next sheetDef
    Exit Sub

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Create tables/settings/named ranges/schema/bootstrap rows", errNumber, errDescription
End Sub

Private Function AddTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal headers As Variant) As ListObject
    Dim index As Long
    Dim tableRange As Range
    Dim lo As ListObject
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    For index = LBound(headers) To UBound(headers)
        ws.Cells(1, index + 1).Value2 = CStr(headers(index))
        ws.Cells(2, index + 1).Value2 = ""
    Next index

    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) + 1))
    Set lo = ws.ListObjects.Add(XL_SRC_RANGE, tableRange, , XL_YES)
    lo.Name = tableName
    lo.TableStyle = "TableStyleMedium2"
    If lo.ListRows.Count > 0 Then lo.ListRows(1).Delete
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).WrapText = True
    ws.Rows(1).RowHeight = 32
    tableRange.EntireColumn.AutoFit
    Set AddTable = lo
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Create table '" & tableName & "' on '" & ws.Name & "'", errNumber, errDescription
End Function

Private Function AddTableRow(ByVal lo As ListObject, ByVal values As Variant) As ListRow
    Dim lr As ListRow
    Dim index As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    Set lr = lo.ListRows.Add
    For index = LBound(values) To UBound(values)
        lr.Range.Cells(1, index + 1).Value2 = values(index)
    Next index
    Set AddTableRow = lr
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Add row to table '" & lo.Name & "'", errNumber, errDescription
End Function

Private Function JoinListColumns(ByVal lo As ListObject) As String
    Dim lc As ListColumn
    Dim result As String
    On Error GoTo Failed
    For Each lc In lo.ListColumns
        If Len(result) > 0 Then result = result & "|"
        result = result & lc.Name
    Next lc
    JoinListColumns = result
    Exit Function
Failed:
    RaiseLoggedError "Read columns from table '" & lo.Name & "'", Err.Number, Err.Description
End Function

Private Function LogSafeValue(ByVal value As Variant) As String
    Dim result As String
    result = CStr(value)
    result = Replace(result, vbCr, "<CR>")
    result = Replace(result, vbLf, "<LF>")
    LogSafeValue = result
End Function

Private Sub FormatWorkbook(ByVal target As Workbook)
    Dim welcome As Worksheet
    Dim dashboard As Worksheet
    Dim vendorsView As Worksheet
    Dim assessmentsView As Worksheet
    Dim labels As Variant
    Dim headers As Variant
    Dim index As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    Set welcome = target.Worksheets("WELCOME")
    Set dashboard = target.Worksheets("Dashboard")
    Set vendorsView = target.Worksheets("Vendors")
    Set assessmentsView = target.Worksheets("Assessments")

    welcome.Cells.Clear
    welcome.Range("B2:H2").Merge
    welcome.Range("B2").Value2 = "Third-Party Assessment Tracker"
    welcome.Range("B2").Font.Size = 24
    welcome.Range("B2").Font.Bold = True
    welcome.Range("B4:H6").Merge
    welcome.Range("B4").Value2 = "If you can read this but buttons do not work, macros are disabled. Close the file, confirm the publisher/signature or approved Trusted Location with IT, then reopen. New Outlook is not supported."
    welcome.Range("B4").WrapText = True
    welcome.Range("B8").Value2 = "Windows identity: checked when macros run"
    welcome.Range("B9").Value2 = "Access: fail-closed until startup checks pass"
    welcome.Range("B10:H10").Merge
    welcome.Range("B:H").ColumnWidth = 18
    welcome.Rows(4).RowHeight = 72

    dashboard.Range("A1:H1").Merge
    dashboard.Range("A1").Value2 = "Third-Party Assessment Dashboard"
    dashboard.Range("A1").Font.Size = 20
    dashboard.Range("A1").Font.Bold = True
    labels = Array("Active assessments", "", "Completed", "", "Due", "", "Overdue", "")
    For index = LBound(labels) To UBound(labels)
        dashboard.Cells(3, index + 1).Value2 = labels(index)
    Next index
    dashboard.Range("A4:H5").Interior.Color = 15921906
    dashboard.Range("B4,D4,F4,H4").Font.Size = 20
    headers = Array("Assessment ID", "Vendor", "Cycle", "Owner", "Owner Email", "Status", "Next Due", "Days Overdue")
    For index = LBound(headers) To UBound(headers)
        dashboard.Cells(12, index + 1).Value2 = headers(index)
    Next index
    dashboard.Range("A12:H12").Font.Bold = True
    dashboard.Range("A:H").ColumnWidth = 18
    dashboard.Range("B:B").ColumnWidth = 28
    dashboard.Range("E:E").ColumnWidth = 30
    dashboard.Range("J:K").EntireColumn.Hidden = True

    AddButton dashboard, "mut_AddVendor", "Add vendor", "modDashboard.OpenVendorForm", 20, 155
    AddButton dashboard, "mut_AddAssessment", "Add assessment", "modDashboard.OpenAssessmentForm", 145, 155
    AddButton dashboard, "mut_Import", "Import", "modImport.StartImport", 270, 155
    AddButton dashboard, "mut_Reconcile", "Reconcile sent", "modEmail.Reconcile", 395, 155
    AddButton dashboard, "mut_BatchDue", "Prepare all due", "modFollowup.PrepareAllDue", 520, 155
    AddButton dashboard, "mgr_Users", "Users", "modUsers.ShowUsersForm", 645, 155
    AddButton dashboard, "mgr_Settings", "Settings", "modDashboard.OpenSettingsForm", 770, 155

    vendorsView.Range("A1:H1").Merge
    vendorsView.Range("A1").Value2 = "Active Vendors"
    vendorsView.Range("A1").Font.Size = 18
    vendorsView.Range("A2").Value2 = "Double-click a row to edit. Business-owner changes affect defaults only, not existing assessment snapshots."
    vendorsView.Range("A2:H2").Merge
    vendorsView.Range("A:H").ColumnWidth = 18
    vendorsView.Range("B:B").ColumnWidth = 28
    vendorsView.Range("H:H").ColumnWidth = 36

    assessmentsView.Range("A1:N1").Merge
    assessmentsView.Range("A1").Value2 = "Active Assessments"
    assessmentsView.Range("A1").Font.Size = 18
    assessmentsView.Range("A2").Value2 = "Double-click a row to edit cycle/owner snapshot/details. Use Status for all status/date changes."
    assessmentsView.Range("A2:N2").Merge
    assessmentsView.Range("A:N").ColumnWidth = 16
    assessmentsView.Range("C:C").ColumnWidth = 26
    assessmentsView.Range("F:F").ColumnWidth = 28
    AddButton assessmentsView, "mut_Status", "Set status", "modDashboard.OpenStatusForm", 20, 55
    AddButton assessmentsView, "mut_Initial", "Prepare initial", "modDashboard.PrepareSelectedInitial", 145, 55
    AddButton assessmentsView, "mut_Followup", "Prepare follow-up", "modDashboard.PrepareSelectedFollowup", 270, 55
    AddButton assessmentsView, "mut_Meeting", "Log meeting", "modDashboard.LogSelectedMeeting", 395, 55

    target.Activate
    welcome.Activate
    ActiveWindow.DisplayGridlines = False
    dashboard.Activate
    ActiveWindow.SplitRow = 12
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    vendorsView.Activate
    ActiveWindow.SplitRow = 4
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    assessmentsView.Activate
    ActiveWindow.SplitRow = 4
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    welcome.Activate
    LogMessage "PASS", "Formatted WELCOME, Dashboard, Vendors, and Assessments views."
    Exit Sub

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Format user-facing worksheets", errNumber, errDescription
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal shapeName As String, _
                      ByVal captionText As String, ByVal macroName As String, _
                      ByVal leftPosition As Double, ByVal topPosition As Double, _
                      Optional ByVal shapeWidth As Double = 118, _
                      Optional ByVal shapeHeight As Double = 28)
    Dim shape As Object
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    Set shape = ws.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, leftPosition, topPosition, shapeWidth, shapeHeight)
    shape.Name = shapeName
    shape.TextFrame2.TextRange.Text = captionText
    shape.TextFrame2.TextRange.Font.Size = 10
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    shape.Fill.ForeColor.RGB = 7626838
    shape.Line.Visible = MSO_FALSE
    shape.OnAction = macroName
    LogMessage "PASS", "Created button '" & shapeName & "' on '" & ws.Name & "' -> " & macroName
    Exit Sub

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Create button '" & shapeName & "' on '" & ws.Name & "'", errNumber, errDescription
End Sub

Private Sub SetWorksheetVisibility(ByVal target As Workbook)
    Dim names As Variant
    Dim item As Variant
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    target.Worksheets("ImportStaging").Visible = XL_SHEET_HIDDEN
    LogMessage "PASS", "Set ImportStaging visibility to Hidden."

    names = Array("Data_Vendors", "Data_Assessments", "Data_EmailEvents", "Data_Users", _
                  "Data_Audit", "Settings", "Schema")
    For Each item In names
        target.Worksheets(CStr(item)).Visible = XL_SHEET_VERY_HIDDEN
        LogMessage "PASS", "Set " & CStr(item) & " visibility to VeryHidden."
    Next item
    Exit Sub

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Set worksheet visibility", errNumber, errDescription
End Sub

Private Sub AddReferenceIfMissing(ByVal project As Object, ByVal guidText As String, _
                                  ByVal majorVersion As Long, ByVal minorVersion As Long, _
                                  ByVal labelText As String)
    Dim reference As Object
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Warning
    For Each reference In project.References
        If StrComp(reference.Guid, guidText, vbTextCompare) = 0 Then
            LogMessage "PASS", "VBA reference already present: " & labelText & " (" & guidText & ")"
            Exit Sub
        End If
    Next reference

    project.References.AddFromGuid guidText, majorVersion, minorVersion
    LogMessage "PASS", "Added VBA reference: " & labelText & " (" & guidText & ")"
    Exit Sub

Warning:
    errNumber = Err.Number
    errDescription = Err.Description
    LogMessage "WARN", "Could not add optional VBA reference '" & labelText & "' (" & guidText & _
               "). Continuing because runtime code is late-bound. Err.Number=" & CStr(errNumber) & _
               " | Err.Description=" & errDescription
    Err.Clear
End Sub

Private Function ImportAllCode(ByVal project As Object, ByVal sourceDirectory As String) As Boolean
    Dim modules As Variant
    Dim forms As Variant
    Dim documents As Variant
    Dim item As Variant
    Dim pair As Variant
    Dim allOK As Boolean
    Dim filePath As String
    Dim tempFolder As String
    Dim tempFrm As String
    Dim tempErrNumber As Long
    Dim tempErrDescription As String

    On Error GoTo Catastrophic
    allOK = True
    modules = RequiredModules()
    forms = RequiredForms()
    documents = DocumentComponents()

    For Each item In modules
        filePath = sourceDirectory & Application.PathSeparator & CStr(item) & ".bas"
        If Not ImportStandardModule(project, CStr(item), filePath) Then allOK = False
    Next item

    On Error Resume Next
    tempFolder = CreateUniqueTempFolder()
    tempErrNumber = Err.Number
    tempErrDescription = Err.Description
    Err.Clear
    On Error GoTo Catastrophic

    If Len(tempFolder) = 0 Then
        allOK = False
        For Each item In forms
            LogError "Import form '" & CStr(item) & "' using isolated .frm-only copy", _
                     tempErrNumber, "No isolated TEMP folder was available: " & tempErrDescription
        Next item
    Else
        LogMessage "INFO", "Form import staging folder: " & tempFolder
        For Each item In forms
            filePath = sourceDirectory & Application.PathSeparator & CStr(item) & ".frm"
            tempFrm = tempFolder & Application.PathSeparator & CStr(item) & ".frm"
            If Not ImportFormWithoutFrx(project, CStr(item), filePath, tempFrm) Then allOK = False
        Next item
        CleanupTempFolder tempFolder
    End If

    For Each pair In documents
        filePath = sourceDirectory & Application.PathSeparator & CStr(pair(1))
        If Not InstallDocumentCode(project, CStr(pair(0)), filePath) Then allOK = False
    Next pair

    ImportAllCode = allOK
    Exit Function

Catastrophic:
    LogError "Import all VBA source files", Err.Number, Err.Description
    If Len(tempFolder) > 0 Then CleanupTempFolder tempFolder
    ImportAllCode = False
    Err.Clear
End Function

Private Function ImportStandardModule(ByVal project As Object, ByVal moduleName As String, _
                                      ByVal filePath As String) As Boolean
    Dim component As Object
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    Set component = project.VBComponents.Import(filePath)
    LogMessage "PASS", "Imported standard module '" & moduleName & "' from " & filePath & _
               " as component '" & component.Name & "'."
    ImportStandardModule = True
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    LogError "Import standard module '" & moduleName & "' from " & filePath, errNumber, errDescription
    ImportStandardModule = False
    Err.Clear
End Function

Private Function CreateUniqueTempFolder() As String
    Dim tempRoot As String
    Dim candidate As String
    Dim suffix As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    tempRoot = Environ$("TEMP")
    If Len(tempRoot) = 0 Then tempRoot = Environ$("TMP")
    If Len(tempRoot) = 0 Or Not FolderExists(tempRoot) Then
        Err.Raise vbObjectError + 4150, "CreateUniqueTempFolder", "Neither TEMP nor TMP resolves to an existing folder."
    End If

    candidate = tempRoot & Application.PathSeparator & "TPAT_FormImport_" & _
                Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 100))
    Do While FolderExists(candidate)
        suffix = suffix + 1
        candidate = candidate & "_" & CStr(suffix)
    Loop
    MkDir candidate
    CreateUniqueTempFolder = candidate
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    RaiseLoggedError "Create isolated TEMP folder for .frm-only imports", errNumber, errDescription
End Function

Private Function ImportFormWithoutFrx(ByVal project As Object, ByVal formName As String, _
                                      ByVal sourceFrm As String, ByVal tempFrm As String) As Boolean
    Dim component As Object
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    If FileExists(tempFrm) Then Kill tempFrm
    FileCopy sourceFrm, tempFrm
    LogMessage "INFO", "Copied only " & formName & ".frm to isolated TEMP folder; no .frx was copied or placed beside it."
    Set component = project.VBComponents.Import(tempFrm)
    LogMessage "PASS", "Imported form '" & formName & "' from temporary .frm-only copy as component '" & component.Name & "'."
    Kill tempFrm
    LogMessage "INFO", "Deleted temporary form copy: " & tempFrm
    ImportFormWithoutFrx = True
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    LogError "Import form '" & formName & "' using isolated .frm-only copy", errNumber, errDescription
    On Error Resume Next
    If FileExists(tempFrm) Then Kill tempFrm
    On Error GoTo 0
    ImportFormWithoutFrx = False
    Err.Clear
End Function

Private Sub CleanupTempFolder(ByVal tempFolder As String)
    Dim errNumber As Long
    Dim errDescription As String
    If Len(tempFolder) = 0 Then Exit Sub
    On Error GoTo Warning
    RmDir tempFolder
    LogMessage "PASS", "Removed isolated form-import TEMP folder: " & tempFolder
    Exit Sub
Warning:
    errNumber = Err.Number
    errDescription = Err.Description
    LogMessage "WARN", "Could not remove form-import TEMP folder: " & tempFolder & _
               " | Err.Number=" & CStr(errNumber) & " | Err.Description=" & errDescription
    Err.Clear
End Sub

Private Function InstallDocumentCode(ByVal project As Object, ByVal componentName As String, _
                                     ByVal sourcePath As String) As Boolean
    Dim component As Object
    Dim codeModule As Object
    Dim codeBody As String
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    codeBody = ReadDocumentCodeBody(sourcePath)
    Set component = project.VBComponents.Item(componentName)
    Set codeModule = component.CodeModule
    If codeModule.CountOfLines > 0 Then codeModule.DeleteLines 1, codeModule.CountOfLines
    codeModule.AddFromString codeBody
    LogMessage "PASS", "Installed document-module code into existing component '" & componentName & _
               "' from " & sourcePath & "."
    InstallDocumentCode = True
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    LogError "Install document code into '" & componentName & "' from " & sourcePath, errNumber, errDescription
    InstallDocumentCode = False
    Err.Clear
End Function

Private Function ReadDocumentCodeBody(ByVal sourcePath As String) As String
    Dim fileNumber As Integer
    Dim currentLine As String
    Dim result As String
    Dim foundMarker As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    fileNumber = FreeFile
    Open sourcePath For Input As #fileNumber
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, currentLine
        If Not foundMarker Then
            If Trim$(currentLine) = "Option Explicit" Then foundMarker = True
        End If
        If foundMarker Then
            If Len(result) > 0 Then result = result & vbCrLf
            result = result & currentLine
        End If
    Loop
    Close #fileNumber
    fileNumber = 0

    If Not foundMarker Then
        Err.Raise vbObjectError + 4160, "ReadDocumentCodeBody", _
            "No Option Explicit marker found in document source: " & sourcePath
    End If
    ReadDocumentCodeBody = result
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    On Error GoTo 0
    Err.Raise errNumber, "ReadDocumentCodeBody", errDescription
End Function

Private Function InstallSelfTestReportShim(ByVal project As Object) As Boolean
    Dim component As Object
    Dim codeModule As Object
    Dim shimText As String
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    Set component = project.VBComponents.Item("modEnv")
    Set codeModule = component.CodeModule
    shimText = vbCrLf & _
        "Public Function BootstrapStructuralSelfTestReport() As String" & vbCrLf & _
        "    Dim report As String" & vbCrLf & _
        "    Dim valid As Boolean" & vbCrLf & _
        "    valid = RunChecks(report, True)" & vbCrLf & _
        "    If valid Then" & vbCrLf & _
        "        BootstrapStructuralSelfTestReport = ""Structural check PASS."" & vbCrLf & SupportInfo()" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        BootstrapStructuralSelfTestReport = ""Structural check FAIL:"" & vbCrLf & report" & vbCrLf & _
        "    End If" & vbCrLf & _
        "End Function"
    codeModule.AddFromString shimText
    LogMessage "PASS", "Installed modEnv.BootstrapStructuralSelfTestReport diagnostic shim. It calls the same private RunChecks routine and exposes the full report text for this build log."
    InstallSelfTestReportShim = True
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    LogError "Install modEnv full-report diagnostic shim", errNumber, errDescription
    InstallSelfTestReportShim = False
    Err.Clear
End Function

Private Function RunStructuralSelfTest(ByVal target As Workbook) As Boolean
    Dim escapedBookName As String
    Dim macroPrefix As String
    Dim testResult As Variant
    Dim reportResult As Variant
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed
    escapedBookName = Replace(target.Name, "'", "''")
    macroPrefix = "'" & escapedBookName & "'!"
    LogMessage "INFO", "Calling structuralOnly=True; Classic Outlook/MAPI and other runtime-only checks are intentionally not required by this build gate."

    testResult = Application.Run(macroPrefix & "modEnv.SelfTest", False, True)
    RunStructuralSelfTest = CBool(testResult)
    LogMessage IIf(RunStructuralSelfTest, "PASS", "FAIL"), _
               "modEnv.SelfTest(False, True) returned: " & CStr(CBool(testResult))

    reportResult = Application.Run(macroPrefix & "modEnv.BootstrapStructuralSelfTestReport")
    LogBlock IIf(RunStructuralSelfTest, "PASS", "FAIL"), _
             "Full modEnv structural self-test report text:", CStr(reportResult)
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    LogError "Run target modEnv.SelfTest(False, True) and capture full report", errNumber, errDescription
    RunStructuralSelfTest = False
    Err.Clear
End Function
