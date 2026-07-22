Attribute VB_Name = "modFollowup"
Option Explicit

Public Function NextFollowupDue(ByVal assessmentID As Long) As Variant
    Dim assessments As ListObject
    Dim lr As ListRow
    Dim lastSent As Variant
    Set assessments = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    Set lr = modUtil.FindLongRow(assessments, "AssessmentID", assessmentID)
    If lr Is Nothing Then NextFollowupDue = Empty: Exit Function
    If modUtil.IsTerminalStatus(CStr(modUtil.RowValue(lr, assessments, "Status"))) Then NextFollowupDue = Empty: Exit Function
    If Not modEmail.HasConfirmedInitial(assessmentID) Then NextFollowupDue = Empty: Exit Function
    lastSent = modEmail.LastSentDate(assessmentID)
    If IsDate(lastSent) Then NextFollowupDue = DateAdd("d", modSettings.GetLong("FollowupDays", 7), CDate(lastSent)) Else NextFollowupDue = Empty
End Function

Public Function IsDue(ByVal assessmentID As Long) As Boolean
    Dim dueValue As Variant
    dueValue = NextFollowupDue(assessmentID)
    IsDue = (IsDate(dueValue) And CDate(dueValue) <= Date)
End Function

Public Sub PrepareAllDue(Optional ByVal includeNeverEmailed As Boolean = False, Optional ByVal stopOnError As Boolean = False)
    Dim assessments As ListObject
    Dim lr As ListRow
    Dim ids() As Long
    Dim kinds() As String
    Dim labels() As String
    Dim count As Long
    Dim cap As Long
    Dim assessmentID As Long
    Dim dueValue As Variant
    Dim prepared As Long
    Dim failed As Long
    Dim i As Long
    Dim summary As String
    If Not modIdentity.RequireMutation() Then Exit Sub
    cap = modSettings.GetLong("BatchCap", 50)
    Set assessments = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    For Each lr In assessments.ListRows
        If Not modUtil.ToBoolean(modUtil.RowValue(lr, assessments, "Archived"), False) And Not modUtil.IsTerminalStatus(CStr(modUtil.RowValue(lr, assessments, "Status"))) Then
            assessmentID = CLng(modUtil.RowValue(lr, assessments, "AssessmentID"))
            dueValue = NextFollowupDue(assessmentID)
            If IsDate(dueValue) And CDate(dueValue) <= Date Then
                AddCandidate ids, kinds, labels, count, assessmentID, "Followup", CStr(modUtil.RowValue(lr, assessments, "BusinessOwnerEmail")) & " | " & modAssessments.VendorNameForID(CLng(modUtil.RowValue(lr, assessments, "VendorID"))) & " | " & CStr(modUtil.RowValue(lr, assessments, "Cycle"))
            ElseIf includeNeverEmailed And Not modEmail.HasConfirmedInitial(assessmentID) And Not modEmail.HasActiveEvent(assessmentID, "Initial") Then
                AddCandidate ids, kinds, labels, count, assessmentID, "Initial", CStr(modUtil.RowValue(lr, assessments, "BusinessOwnerEmail")) & " | " & modAssessments.VendorNameForID(CLng(modUtil.RowValue(lr, assessments, "VendorID"))) & " | " & CStr(modUtil.RowValue(lr, assessments, "Cycle"))
            End If
        End If
    Next lr
    If count = 0 Then MsgBox "No eligible assessments were found.", vbInformation, modUtil.APP_NAME: Exit Sub
    SortCandidates labels, ids, kinds, count
    If count > cap Then count = cap
    summary = "The first " & count & " item(s), grouped by business owner, will be prepared." & vbCrLf & "Mode: " & modSettings.GetSetting("EmailMode", "Draft") & vbCrLf & "Batch cap: " & cap
    frmBatchPreview.LoadPreview labels, count, summary
    frmBatchPreview.Show
    If Not frmBatchPreview.Confirmed Then Exit Sub
    If StrComp(modSettings.GetSetting("EmailMode", "Draft"), "Direct", vbTextCompare) = 0 Then
        If Not modIdentity.RequireRole("Manager") Then Exit Sub
        If MsgBox("DIRECT SEND will call Outlook .Send for " & count & " recipient(s). Items are recorded as Queued, not Sent, until reconciliation. Continue?", vbCritical + vbYesNo + vbDefaultButton2, modUtil.APP_NAME) <> vbYes Then Exit Sub
    End If
    For i = 1 To count
        If modEmail.PrepareEmail(ids(i), kinds(i), False, True) > 0 Then
            prepared = prepared + 1
        Else
            failed = failed + 1
            If stopOnError Then Exit For
        End If
        DoEvents
    Next i
    MsgBox "Prepared/queued: " & prepared & vbCrLf & "Failed/skipped: " & failed, vbInformation, modUtil.APP_NAME
End Sub

Private Sub AddCandidate(ByRef ids() As Long, ByRef kinds() As String, ByRef labels() As String, ByRef count As Long, ByVal assessmentID As Long, ByVal kind As String, ByVal label As String)
    count = count + 1
    ReDim Preserve ids(1 To count)
    ReDim Preserve kinds(1 To count)
    ReDim Preserve labels(1 To count)
    ids(count) = assessmentID
    kinds(count) = kind
    labels(count) = label
End Sub

Private Sub SortCandidates(ByRef labels() As String, ByRef ids() As Long, ByRef kinds() As String, ByVal count As Long)
    Dim i As Long, j As Long, tempLabel As String, tempID As Long, tempKind As String
    For i = 1 To count - 1
        For j = i + 1 To count
            If StrComp(labels(j), labels(i), vbTextCompare) < 0 Then
                tempLabel = labels(i): labels(i) = labels(j): labels(j) = tempLabel
                tempID = ids(i): ids(i) = ids(j): ids(j) = tempID
                tempKind = kinds(i): kinds(i) = kinds(j): kinds(j) = tempKind
            End If
        Next j
    Next i
End Sub
