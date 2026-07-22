Attribute VB_Name = "modImport"
Option Explicit

Private Const MSO_AUTOMATION_SECURITY_FORCE_DISABLE As Long = 3
Private Const XL_CALCULATION_MANUAL As Long = -4135
Private mBusy As Boolean

Public Sub StartImport()
    Dim selectedFile As Variant
    Dim importMode As String
    If Not modIdentity.RequireMutation() Then Exit Sub
    selectedFile = Application.GetOpenFilename("Supported files (*.csv;*.xlsx),*.csv;*.xlsx", , "Choose vendor import file")
    If VarType(selectedFile) = vbBoolean Then Exit Sub
    importMode = UCase$(Trim$(InputBox("Enter INSERT or UPDATE.", modUtil.APP_NAME, "INSERT")))
    If importMode <> "INSERT" And importMode <> "UPDATE" Then MsgBox "Import cancelled: mode must be INSERT or UPDATE.", vbExclamation, modUtil.APP_NAME: Exit Sub
    StageFile CStr(selectedFile), importMode
End Sub

Public Sub StageFile(ByVal filePath As String, ByVal importMode As String)
    Dim extension As String
    Dim rows As Collection
    If mBusy Then Exit Sub
    If Not modIdentity.RequireMutation() Then Exit Sub
    extension = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))
    On Error GoTo Failed
    mBusy = True
    If extension = "csv" Then
        Set rows = ParseCsvUtf8(filePath)
    ElseIf extension = "xlsx" Then
        Set rows = ReadXlsxTextFirst(filePath)
    Else
        Err.Raise vbObjectError + 1800, , "Only .csv and .xlsx files are accepted."
    End If
    StageRows rows, importMode, filePath
Done:
    mBusy = False
    Exit Sub
Failed:
    modUtil.ShowError "Import", Err.Description
    Resume Done
End Sub

Private Function ParseCsvUtf8(ByVal filePath As String) As Collection
    Dim stream As Object
    Dim text As String
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    text = stream.ReadText
    stream.Close
    If Left$(text, 1) = ChrW$(&HFEFF) Then text = Mid$(text, 2)
    Set ParseCsvUtf8 = ParseCsvText(text)
End Function

Public Function ParseCsvText(ByVal text As String) As Collection
    Dim rows As New Collection
    Dim fields As Collection
    Dim fieldValue As String
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Set fields = New Collection
    i = 1
    Do While i <= Len(text)
        ch = Mid$(text, i, 1)
        If inQuotes Then
            If ch = """" Then
                If i < Len(text) And Mid$(text, i + 1, 1) = """" Then
                    fieldValue = fieldValue & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                fieldValue = fieldValue & ch
            End If
        Else
            Select Case ch
                Case """": inQuotes = True
                Case ",": fields.Add fieldValue: fieldValue = vbNullString
                Case vbCr
                    If i < Len(text) And Mid$(text, i + 1, 1) = vbLf Then i = i + 1
                    fields.Add fieldValue: fieldValue = vbNullString
                    rows.Add CollectionToArray(fields): Set fields = New Collection
                Case vbLf
                    fields.Add fieldValue: fieldValue = vbNullString
                    rows.Add CollectionToArray(fields): Set fields = New Collection
                Case Else: fieldValue = fieldValue & ch
            End Select
        End If
        i = i + 1
    Loop
    If inQuotes Then Err.Raise vbObjectError + 1801, , "CSV contains an unterminated quoted field."
    If fields.count > 0 Or Len(fieldValue) > 0 Then fields.Add fieldValue: rows.Add CollectionToArray(fields)
    Set ParseCsvText = rows
End Function

Private Function CollectionToArray(ByVal values As Collection) As Variant
    Dim result() As String
    Dim i As Long
    ReDim result(1 To values.count)
    For i = 1 To values.count: result(i) = CStr(values(i)): Next i
    CollectionToArray = result
End Function

Private Function ReadXlsxTextFirst(ByVal filePath As String) As Collection
    Dim oldSecurity As Long
    Dim oldEvents As Boolean
    Dim oldCalc As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim used As Range
    Dim rows As New Collection
    Dim result() As String
    Dim r As Long, c As Long
    Dim errorNumber As Long
    Dim errorDescription As String
    On Error GoTo Failed
    oldSecurity = Application.AutomationSecurity
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    Application.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    Application.EnableEvents = False
    Application.Calculation = XL_CALCULATION_MANUAL
    Set wb = Application.Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    Set ws = wb.Worksheets(1)
    Set used = ws.UsedRange
    For r = 1 To used.Rows.count
        ReDim result(1 To used.Columns.count)
        For c = 1 To used.Columns.count
            If used.Cells(r, c).HasFormula Then Err.Raise vbObjectError + 1802, , "Formula cell rejected at " & used.Cells(r, c).Address(False, False) & ". Save values-only data and retry."
            result(c) = CStr(used.Cells(r, c).Value2)
        Next c
        rows.Add result
    Next r
    wb.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    Set ReadXlsxTextFirst = rows
    Exit Function
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    On Error GoTo 0
    Err.Raise errorNumber, "modImport.ReadXlsxTextFirst", errorDescription
End Function

Private Sub StageRows(ByVal rows As Collection, ByVal importMode As String, ByVal sourceFile As String)
    Dim headers As Object
    Dim requiredHeaders As Variant
    Dim rowData As Variant
    Dim lo As ListObject
    Dim lr As ListRow
    Dim i As Long
    Dim validCount As Long
    Dim errorCount As Long
    Dim errorText As String
    Dim actionText As String
    If rows.count < 2 Then Err.Raise vbObjectError + 1803, , "Import file has no data rows."
    Set headers = HeaderMap(rows(1))
    requiredHeaders = Array("vendorname", "businessownername", "businessowneremail", "vendorcontactperson", "vendorcontactemail", "vendorcontactphone", "notes", "cycle")
    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        If Not headers.Exists(requiredHeaders(i)) Then Err.Raise vbObjectError + 1804, , "Missing header: " & requiredHeaders(i)
    Next i
    Set lo = modUtil.GetTable(modUtil.TBL_IMPORT)
    Do While lo.ListRows.count > 0: lo.ListRows(1).Delete: Loop
    For i = 2 To rows.count
        rowData = rows(i)
        errorText = ValidateImportRow(rowData, headers, importMode, actionText)
        Set lr = lo.ListRows.Add
        modUtil.SetRowValue lr, lo, "SourceRow", i
        modUtil.SetRowValue lr, lo, "ImportMode", importMode
        modUtil.SetRowValue lr, lo, "Action", actionText
        modUtil.SetRowValue lr, lo, "Valid", (Len(errorText) = 0)
        modUtil.SetRowValue lr, lo, "Errors", errorText
        WriteStageText lr, lo, "VendorName", Field(rowData, headers, "vendorname")
        WriteStageText lr, lo, "BusinessOwnerName", Field(rowData, headers, "businessownername")
        WriteStageText lr, lo, "BusinessOwnerEmail", Field(rowData, headers, "businessowneremail")
        WriteStageText lr, lo, "VendorContactPerson", Field(rowData, headers, "vendorcontactperson")
        WriteStageText lr, lo, "VendorContactEmail", Field(rowData, headers, "vendorcontactemail")
        WriteStageText lr, lo, "VendorContactPhone", Field(rowData, headers, "vendorcontactphone")
        WriteStageText lr, lo, "Notes", Field(rowData, headers, "notes")
        WriteStageText lr, lo, "Cycle", Field(rowData, headers, "cycle")
        If Len(errorText) = 0 Then validCount = validCount + 1 Else errorCount = errorCount + 1
    Next i
    ThisWorkbook.Worksheets("ImportStaging").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("ImportStaging").Activate
    mBusy = False
    frmImportReview.LoadReview validCount, errorCount, sourceFile, importMode
    frmImportReview.Show
End Sub

Private Function HeaderMap(ByVal headerRow As Variant) As Object
    Dim map As Object
    Dim i As Long
    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = 1
    For i = LBound(headerRow) To UBound(headerRow): map(LCase$(Trim$(CStr(headerRow(i))))) = i: Next i
    Set HeaderMap = map
End Function

Private Function Field(ByVal rowData As Variant, ByVal headers As Object, ByVal name As String) As String
    Dim idx As Long
    idx = CLng(headers(name))
    If idx <= UBound(rowData) Then Field = CStr(rowData(idx))
End Function

Private Function ValidateImportRow(ByVal rowData As Variant, ByVal headers As Object, ByVal importMode As String, ByRef actionText As String) As String
    Dim vendorName As String, ownerName As String, ownerEmail As String, contactEmail As String, cycle As String
    Dim vendorID As Long
    vendorName = Trim$(Field(rowData, headers, "vendorname"))
    ownerName = Trim$(Field(rowData, headers, "businessownername"))
    ownerEmail = Trim$(Field(rowData, headers, "businessowneremail"))
    contactEmail = Trim$(Field(rowData, headers, "vendorcontactemail"))
    cycle = Trim$(Field(rowData, headers, "cycle"))
    If Len(vendorName) = 0 Then AppendError ValidateImportRow, "VendorName required"
    If Len(ownerName) = 0 Then AppendError ValidateImportRow, "BusinessOwnerName required"
    If Not modUtil.IsEmailAddress(ownerEmail) Then AppendError ValidateImportRow, "BusinessOwnerEmail invalid"
    If Len(contactEmail) > 0 And Not modUtil.IsEmailAddress(contactEmail) Then AppendError ValidateImportRow, "VendorContactEmail invalid"
    vendorID = modVendors.FindVendorByBusinessKey(vendorName, ownerEmail)
    If importMode = "UPDATE" Then
        If vendorID = 0 Then
            AppendError ValidateImportRow, "Update key not found (VendorName + BusinessOwnerEmail)"
            actionText = "Error"
        Else
            actionText = "Update VendorID " & vendorID
        End If
    Else
        actionText = "Insert"
    End If
    If vendorID > 0 And Len(cycle) > 0 And modAssessments.AssessmentPairExists(vendorID, cycle) Then AppendError ValidateImportRow, "Assessment cycle already exists for matched vendor"
End Function

Private Sub AppendError(ByRef target As String, ByVal messageText As String)
    If Len(target) > 0 Then target = target & "; "
    target = target & messageText
End Sub

Private Sub WriteStageText(ByVal lr As ListRow, ByVal lo As ListObject, ByVal columnName As String, ByVal value As String)
    modUtil.SetRowValue lr, lo, columnName, modUtil.NeutralizeImportedText(value)
End Sub

Public Sub CommitStagedImport()
    Dim lo As ListObject
    Dim lr As ListRow
    Dim mode As String
    Dim vendorID As Long
    Dim createdAssessment As Long
    Dim vendors As ListObject
    Dim vendorRow As ListRow
    Dim targetVendorName As String
    Dim targetOwnerName As String
    Dim targetOwnerEmail As String
    Dim targetContactPerson As String
    Dim targetContactEmail As String
    Dim targetContactPhone As String
    Dim targetNotes As String
    Dim committed As Long
    Dim failed As Long
    If mBusy Then Exit Sub
    If Not modIdentity.RequireMutation() Then Exit Sub
    On Error GoTo Fatal
    modBackup.RequireRiskBackup
    mBusy = True
    Set lo = modUtil.GetTable(modUtil.TBL_IMPORT)
    For Each lr In lo.ListRows
        If modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Valid"), False) Then
            mode = CStr(modUtil.RowValue(lr, lo, "ImportMode"))
            If mode = "UPDATE" Then
                vendorID = modVendors.FindVendorByBusinessKey(CleanStage(modUtil.RowValue(lr, lo, "VendorName")), CleanStage(modUtil.RowValue(lr, lo, "BusinessOwnerEmail")))
                If vendorID > 0 Then
                    Set vendors = modUtil.GetTable(modUtil.TBL_VENDORS)
                    Set vendorRow = modUtil.FindLongRow(vendors, "VendorID", vendorID)
                    targetVendorName = CStr(modUtil.RowValue(vendorRow, vendors, "VendorName"))
                    targetOwnerEmail = CStr(modUtil.RowValue(vendorRow, vendors, "BusinessOwnerEmail"))
                    If frmImportReview.UpdateOwnerName() Then targetOwnerName = CleanStage(modUtil.RowValue(lr, lo, "BusinessOwnerName")) Else targetOwnerName = CStr(modUtil.RowValue(vendorRow, vendors, "BusinessOwnerName"))
                    If frmImportReview.UpdateContactPerson() Then targetContactPerson = CleanStage(modUtil.RowValue(lr, lo, "VendorContactPerson")) Else targetContactPerson = CStr(modUtil.RowValue(vendorRow, vendors, "VendorContactPerson"))
                    If frmImportReview.UpdateContactEmail() Then targetContactEmail = CleanStage(modUtil.RowValue(lr, lo, "VendorContactEmail")) Else targetContactEmail = CStr(modUtil.RowValue(vendorRow, vendors, "VendorContactEmail"))
                    If frmImportReview.UpdateContactPhone() Then targetContactPhone = CleanStage(modUtil.RowValue(lr, lo, "VendorContactPhone")) Else targetContactPhone = CStr(modUtil.RowValue(vendorRow, vendors, "VendorContactPhone"))
                    If frmImportReview.UpdateNotes() Then targetNotes = CleanStage(modUtil.RowValue(lr, lo, "Notes")) Else targetNotes = CStr(modUtil.RowValue(vendorRow, vendors, "Notes"))
                    If Not modVendors.UpdateVendor(vendorID, targetVendorName, targetOwnerName, targetOwnerEmail, targetContactPerson, targetContactEmail, targetContactPhone, targetNotes) Then vendorID = 0
                End If
            Else
                vendorID = modVendors.CreateVendor(CleanStage(modUtil.RowValue(lr, lo, "VendorName")), CleanStage(modUtil.RowValue(lr, lo, "BusinessOwnerName")), CleanStage(modUtil.RowValue(lr, lo, "BusinessOwnerEmail")), CleanStage(modUtil.RowValue(lr, lo, "VendorContactPerson")), CleanStage(modUtil.RowValue(lr, lo, "VendorContactEmail")), CleanStage(modUtil.RowValue(lr, lo, "VendorContactPhone")), CleanStage(modUtil.RowValue(lr, lo, "Notes")), True)
            End If
            If vendorID > 0 And Len(CleanStage(modUtil.RowValue(lr, lo, "Cycle"))) > 0 Then createdAssessment = modAssessments.CreateAssessment(vendorID, CleanStage(modUtil.RowValue(lr, lo, "Cycle")), CleanStage(modUtil.RowValue(lr, lo, "Notes")))
            If vendorID > 0 Then committed = committed + 1 Else failed = failed + 1
        End If
    Next lr
    modAudit.AuditOrRaise "IMPORT_COMMIT", "Import", Format$(Now, "yyyymmddhhnnss"), "Committed=" & committed & "; Failed=" & failed
    mBusy = False
    modDashboard.RefreshDashboard
    MsgBox "Import committed: " & committed & " row(s). Failed during commit: " & failed & ". Invalid staged rows were not committed.", vbInformation, modUtil.APP_NAME
Done:
    mBusy = False
    Exit Sub
Fatal:
    modUtil.ShowError "Import commit", Err.Description
    Resume Done
End Sub

Private Function CleanStage(ByVal value As Variant) As String
    CleanStage = CStr(value)
End Function

Public Function ImportInProgress() As Boolean
    ImportInProgress = mBusy
End Function
