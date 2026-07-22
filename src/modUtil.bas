Attribute VB_Name = "modUtil"
Option Explicit

Public Const APP_NAME As String = "Third-Party Assessment Tracker"
Public Const SCHEMA_VERSION As String = "2.0.0"

Public Const TBL_VENDORS As String = "tblVendors"
Public Const TBL_ASSESSMENTS As String = "tblAssessments"
Public Const TBL_EMAIL_EVENTS As String = "tblEmailEvents"
Public Const TBL_USERS As String = "tblUsers"
Public Const TBL_AUDIT As String = "tblAudit"
Public Const TBL_SETTINGS As String = "tblSettings"
Public Const TBL_IMPORT As String = "tblImportStaging"
Public Const TBL_SCHEMA As String = "tblSchemaManifest"

Private Type TGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef guid As TGuid) As Long
    Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (ByRef guid As TGuid, ByVal buffer As LongPtr, ByVal bufferLength As Long) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef guid As TGuid) As Long
    Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByRef guid As TGuid, ByVal buffer As Long, ByVal bufferLength As Long) As Long
#End If

Public Function GetTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set GetTable = lo
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 1000, "modUtil.GetTable", "Required table not found: " & tableName
End Function

Public Function ColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    On Error GoTo Missing
    ColumnIndex = lo.ListColumns(columnName).Index
    Exit Function
Missing:
    Err.Raise vbObjectError + 1001, "modUtil.ColumnIndex", "Required column '" & columnName & "' is missing from " & lo.Name
End Function

Public Function FindLongRow(ByVal lo As ListObject, ByVal columnName As String, ByVal key As Long) As ListRow
    Dim lr As ListRow
    Dim idx As Long
    idx = ColumnIndex(lo, columnName)
    For Each lr In lo.ListRows
        If Len(Trim$(CStr(lr.Range.Cells(1, idx).Value2))) > 0 Then
            If CLng(lr.Range.Cells(1, idx).Value2) = key Then
                Set FindLongRow = lr
                Exit Function
            End If
        End If
    Next lr
End Function

Public Function FindTextRow(ByVal lo As ListObject, ByVal columnName As String, ByVal key As String) As ListRow
    Dim lr As ListRow
    Dim idx As Long
    idx = ColumnIndex(lo, columnName)
    For Each lr In lo.ListRows
        If StrComp(Trim$(CStr(lr.Range.Cells(1, idx).Value2)), Trim$(key), vbTextCompare) = 0 Then
            Set FindTextRow = lr
            Exit Function
        End If
    Next lr
End Function

Public Function RowValue(ByVal lr As ListRow, ByVal lo As ListObject, ByVal columnName As String) As Variant
    RowValue = lr.Range.Cells(1, ColumnIndex(lo, columnName)).Value
End Function

Public Sub SetRowValue(ByVal lr As ListRow, ByVal lo As ListObject, ByVal columnName As String, ByVal value As Variant)
    lr.Range.Cells(1, ColumnIndex(lo, columnName)).Value = value
End Sub

Public Function NzText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        NzText = vbNullString
    Else
        NzText = Trim$(CStr(value))
    End If
End Function

Public Function ToBoolean(ByVal value As Variant, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim s As String
    If VarType(value) = vbBoolean Then
        ToBoolean = CBool(value)
        Exit Function
    End If
    s = LCase$(Trim$(CStr(value)))
    If s = "true" Or s = "yes" Or s = "1" Then
        ToBoolean = True
    ElseIf s = "false" Or s = "no" Or s = "0" Then
        ToBoolean = False
    Else
        ToBoolean = defaultValue
    End If
End Function

Public Function IsEmailAddress(ByVal emailAddress As String) As Boolean
    Dim atPos As Long
    Dim localPart As String
    Dim domainPart As String
    Dim labels As Variant
    Dim label As Variant
    Dim i As Long
    Dim ch As String
    emailAddress = Trim$(emailAddress)
    If Len(emailAddress) < 5 Or InStr(emailAddress, " ") > 0 Then Exit Function
    atPos = InStrRev(emailAddress, "@")
    If atPos <= 1 Or atPos <> InStr(1, emailAddress, "@") Then Exit Function
    localPart = Left$(emailAddress, atPos - 1)
    domainPart = Mid$(emailAddress, atPos + 1)
    If Len(localPart) > 64 Or Left$(localPart, 1) = "." Or Right$(localPart, 1) = "." Or InStr(localPart, "..") > 0 Then Exit Function
    If InStr(domainPart, ".") <= 1 Or Right$(domainPart, 1) = "." Or InStr(domainPart, "..") > 0 Or Len(domainPart) > 253 Then Exit Function
    For i = 1 To Len(localPart)
        ch = Mid$(localPart, i, 1)
        If AscW(ch) < 33 Or InStr("(),:;<>[]" & Chr$(34), ch) > 0 Then Exit Function
    Next i
    labels = Split(domainPart, ".")
    For Each label In labels
        If Len(CStr(label)) = 0 Or Len(CStr(label)) > 63 Then Exit Function
        If Left$(CStr(label), 1) = "-" Or Right$(CStr(label), 1) = "-" Then Exit Function
        For i = 1 To Len(CStr(label))
            ch = Mid$(CStr(label), i, 1)
            If Not (ch Like "[A-Za-z0-9-]") Then Exit Function
        Next i
    Next label
    IsEmailAddress = True
End Function

Public Function EmailDomain(ByVal emailAddress As String) As String
    Dim p As Long
    p = InStrRev(Trim$(emailAddress), "@")
    If p > 0 Then EmailDomain = LCase$(Trim$(Mid$(emailAddress, p + 1)))
End Function

Public Function IsAllowedInternalEmail(ByVal emailAddress As String) As Boolean
    Dim allowlist As String
    Dim parts() As String
    Dim item As Variant
    Dim domainPart As String
    domainPart = EmailDomain(emailAddress)
    If Len(domainPart) = 0 Then Exit Function
    allowlist = Replace(modSettings.GetSetting("InternalDomainAllowlist", vbNullString), ",", ";")
    If Len(Trim$(allowlist)) = 0 Then Exit Function
    parts = Split(allowlist, ";")
    For Each item In parts
        item = LCase$(Trim$(CStr(item)))
        If Left$(CStr(item), 1) = "@" Then item = Mid$(CStr(item), 2)
        If domainPart = CStr(item) Then
            IsAllowedInternalEmail = True
            Exit Function
        End If
    Next item
End Function

Public Function NeutralizeImportedText(ByVal value As Variant) As String
    Dim s As String
    s = CStr(value)
    If Len(s) > 0 Then
        Select Case Left$(s, 1)
            Case "=", "+", "-", "@", vbTab, vbCr
                s = "'" & s
        End Select
    End If
    NeutralizeImportedText = s
End Function

Public Function StatusRank(ByVal statusName As String) As Long
    Select Case LCase$(Trim$(statusName))
        Case "not started": StatusRank = 0
        Case "in progress": StatusRank = 1
        Case "submitted": StatusRank = 2
        Case "completed": StatusRank = 3
        Case Else: StatusRank = -1
    End Select
End Function

Public Function IsTerminalStatus(ByVal statusName As String) As Boolean
    IsTerminalStatus = (StatusRank(statusName) >= 2)
End Function

Public Function NewGuidToken() As String
    Dim guid As TGuid
    Dim buffer As String
    Dim result As Long
    If CoCreateGuid(guid) <> 0 Then Err.Raise vbObjectError + 1002, "modUtil.NewGuidToken", "Windows could not generate a GUID."
    buffer = String$(39, vbNullChar)
    result = StringFromGUID2(guid, StrPtr(buffer), 39)
    If result = 0 Then Err.Raise vbObjectError + 1003, "modUtil.NewGuidToken", "Windows could not format a GUID."
    NewGuidToken = LCase$(Left$(buffer, result - 1))
End Function

Public Function MaxExistingID(ByVal lo As ListObject, ByVal idColumn As String) As Long
    Dim lr As ListRow
    Dim idx As Long
    Dim candidate As Variant
    idx = ColumnIndex(lo, idColumn)
    For Each lr In lo.ListRows
        candidate = lr.Range.Cells(1, idx).Value2
        If IsNumeric(candidate) Then
            If CDbl(candidate) > MaxExistingID Then MaxExistingID = CLng(candidate)
        End If
    Next lr
End Function

Public Function AllocateID(ByVal entityName As String) As Long
    Dim lo As ListObject
    Dim idColumn As String
    Dim settingName As String
    Dim storedNext As Long
    Dim newID As Long
    Select Case LCase$(entityName)
        Case "vendor"
            Set lo = GetTable(TBL_VENDORS): idColumn = "VendorID": settingName = "NextVendorID"
        Case "assessment"
            Set lo = GetTable(TBL_ASSESSMENTS): idColumn = "AssessmentID": settingName = "NextAssessmentID"
        Case "event"
            Set lo = GetTable(TBL_EMAIL_EVENTS): idColumn = "EventID": settingName = "NextEventID"
        Case Else
            Err.Raise vbObjectError + 1004, "modUtil.AllocateID", "Unknown ID entity: " & entityName
    End Select
    storedNext = modSettings.GetLong(settingName, 1)
    newID = storedNext
    If MaxExistingID(lo, idColumn) + 1 > newID Then newID = MaxExistingID(lo, idColumn) + 1
    If Not FindLongRow(lo, idColumn, newID) Is Nothing Then Err.Raise vbObjectError + 1005, "modUtil.AllocateID", "Allocated ID is already in use."
    modSettings.SetSystemSetting settingName, CStr(newID + 1)
    AllocateID = newID
End Function

Public Function RenderTemplate(ByVal templateText As String, ByVal vendorName As String, ByVal ownerName As String, ByVal cycle As String) As String
    Dim rendered As String
    rendered = templateText
    rendered = Replace(rendered, "{VendorName}", vendorName, 1, -1, vbTextCompare)
    rendered = Replace(rendered, "{OwnerName}", ownerName, 1, -1, vbTextCompare)
    rendered = Replace(rendered, "{OrgName}", modSettings.GetSetting("OrgName", vbNullString), 1, -1, vbTextCompare)
    rendered = Replace(rendered, "{Cycle}", cycle, 1, -1, vbTextCompare)
    If InStr(rendered, "{") > 0 Or InStr(rendered, "}") > 0 Then
        Err.Raise vbObjectError + 1006, "modUtil.RenderTemplate", "Email template contains an unresolved token."
    End If
    RenderTemplate = rendered
End Function

Public Function SafeFileBaseName(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 1 Then SafeFileBaseName = Left$(fileName, p - 1) Else SafeFileBaseName = fileName
End Function

Public Function ValuesEqual(ByVal firstValue As Variant, ByVal secondValue As Variant) As Boolean
    If IsDate(firstValue) And IsDate(secondValue) Then
        ValuesEqual = (CDbl(CDate(firstValue)) = CDbl(CDate(secondValue)))
    Else
        ValuesEqual = (NzText(firstValue) = NzText(secondValue))
    End If
End Function

Public Sub ShowError(ByVal sourceName As String, ByVal messageText As String)
    MsgBox messageText, vbCritical, APP_NAME & " - " & sourceName
End Sub
