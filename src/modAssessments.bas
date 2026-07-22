Attribute VB_Name = "modAssessments"
Option Explicit

Private mBusy As Boolean

Public Function CreateAssessment(ByVal vendorID As Long, ByVal cycle As String, Optional ByVal notes As String = vbNullString, _
    Optional ByVal snapshotOwnerName As String = vbNullString, Optional ByVal snapshotOwnerEmail As String = vbNullString) As Long
    Dim vendors As ListObject
    Dim vendorRow As ListRow
    Dim assessments As ListObject
    Dim lr As ListRow
    Dim assessmentID As Long
    If mBusy Then Exit Function
    If Not modIdentity.RequireMutation() Then Exit Function
    cycle = Trim$(cycle)
    If Len(cycle) = 0 Then MsgBox "Cycle is required.", vbExclamation, modUtil.APP_NAME: Exit Function
    Set vendors = modUtil.GetTable(modUtil.TBL_VENDORS)
    Set vendorRow = modUtil.FindLongRow(vendors, "VendorID", vendorID)
    If vendorRow Is Nothing Then MsgBox "Vendor does not exist.", vbExclamation, modUtil.APP_NAME: Exit Function
    If modUtil.ToBoolean(modUtil.RowValue(vendorRow, vendors, "Archived"), False) Then MsgBox "Archived vendors cannot receive new assessments.", vbExclamation, modUtil.APP_NAME: Exit Function
    If AssessmentPairExists(vendorID, cycle) Then MsgBox "This vendor already has an assessment for that cycle.", vbExclamation, modUtil.APP_NAME: Exit Function
    If Len(snapshotOwnerName) = 0 Then snapshotOwnerName = CStr(modUtil.RowValue(vendorRow, vendors, "BusinessOwnerName"))
    If Len(snapshotOwnerEmail) = 0 Then snapshotOwnerEmail = CStr(modUtil.RowValue(vendorRow, vendors, "BusinessOwnerEmail"))
    If Len(Trim$(snapshotOwnerName)) = 0 Or Not modUtil.IsEmailAddress(snapshotOwnerEmail) Then MsgBox "The assessment owner snapshot is invalid.", vbExclamation, modUtil.APP_NAME: Exit Function
    On Error GoTo Failed
    mBusy = True
    Set assessments = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    assessmentID = modUtil.AllocateID("Assessment")
    Set lr = assessments.ListRows.Add
    modUtil.SetRowValue lr, assessments, "AssessmentID", assessmentID
    modUtil.SetRowValue lr, assessments, "VendorID", vendorID
    modUtil.SetRowValue lr, assessments, "Cycle", modUtil.NeutralizeImportedText(cycle)
    modUtil.SetRowValue lr, assessments, "BusinessOwnerName", modUtil.NeutralizeImportedText(Trim$(snapshotOwnerName))
    modUtil.SetRowValue lr, assessments, "BusinessOwnerEmail", LCase$(Trim$(snapshotOwnerEmail))
    modUtil.SetRowValue lr, assessments, "Status", "Not Started"
    modUtil.SetRowValue lr, assessments, "MeetingsConducted", 0
    modUtil.SetRowValue lr, assessments, "Notes", modUtil.NeutralizeImportedText(notes)
    modUtil.SetRowValue lr, assessments, "CreatedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, assessments, "CreatedOn", Now
    modUtil.SetRowValue lr, assessments, "ModifiedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, assessments, "ModifiedOn", Now
    modUtil.SetRowValue lr, assessments, "Archived", False
    If Not modAudit.AddAudit("CREATE", "Assessment", assessmentID, "VendorID=" & vendorID & "; Cycle=" & cycle & "; owner snapshot persisted") Then
        lr.Delete
        Err.Raise vbObjectError + 1500, , "The assessment was not created because the audit record failed."
    End If
    CreateAssessment = assessmentID
    If Not modImport.ImportInProgress() Then modDashboard.RefreshDashboard
Done:
    mBusy = False
    Exit Function
Failed:
    On Error Resume Next: If Not lr Is Nothing Then lr.Delete
    On Error GoTo 0
    modUtil.ShowError "Assessment", Err.Description
    Resume Done
End Function

Public Function AssessmentPairExists(ByVal vendorID As Long, ByVal cycle As String, Optional ByVal exceptAssessmentID As Long = 0) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    Dim cycleKey As String
    cycleKey = modUtil.NeutralizeImportedText(Trim$(cycle))
    Set lo = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    For Each lr In lo.ListRows
        If CLng(modUtil.RowValue(lr, lo, "AssessmentID")) <> exceptAssessmentID Then
            If CLng(modUtil.RowValue(lr, lo, "VendorID")) = vendorID And StrComp(Trim$(CStr(modUtil.RowValue(lr, lo, "Cycle"))), cycleKey, vbTextCompare) = 0 Then
                AssessmentPairExists = True
                Exit Function
            End If
        End If
    Next lr
End Function

Public Function UpdateAssessmentDetails(ByVal assessmentID As Long, ByVal cycle As String, ByVal ownerName As String, ByVal ownerEmail As String, ByVal notes As String) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    Dim oldValues As Variant
    Dim vendorID As Long
    If mBusy Then Exit Function
    If Not modIdentity.RequireMutation() Then Exit Function
    Set lo = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    Set lr = modUtil.FindLongRow(lo, "AssessmentID", assessmentID)
    If lr Is Nothing Then MsgBox "Assessment not found.", vbExclamation, modUtil.APP_NAME: Exit Function
    vendorID = CLng(modUtil.RowValue(lr, lo, "VendorID"))
    If Len(Trim$(cycle)) = 0 Or Len(Trim$(ownerName)) = 0 Or Not modUtil.IsEmailAddress(ownerEmail) Then MsgBox "Cycle and a valid owner snapshot are required.", vbExclamation, modUtil.APP_NAME: Exit Function
    If AssessmentPairExists(vendorID, cycle, assessmentID) Then MsgBox "This vendor already has an assessment for that cycle.", vbExclamation, modUtil.APP_NAME: Exit Function
    oldValues = lr.Range.Value2
    On Error GoTo Failed
    mBusy = True
    modUtil.SetRowValue lr, lo, "Cycle", modUtil.NeutralizeImportedText(Trim$(cycle))
    modUtil.SetRowValue lr, lo, "BusinessOwnerName", modUtil.NeutralizeImportedText(Trim$(ownerName))
    modUtil.SetRowValue lr, lo, "BusinessOwnerEmail", LCase$(Trim$(ownerEmail))
    modUtil.SetRowValue lr, lo, "Notes", modUtil.NeutralizeImportedText(notes)
    modUtil.SetRowValue lr, lo, "ModifiedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, lo, "ModifiedOn", Now
    If Not modAudit.AddAudit("UPDATE", "Assessment", assessmentID, "Cycle/owner snapshot/details updated; status dates unchanged.") Then lr.Range.Value2 = oldValues: Err.Raise vbObjectError + 1501, , "Audit failed; assessment unchanged."
    UpdateAssessmentDetails = True
Done:
    mBusy = False
    Exit Function
Failed:
    On Error Resume Next: lr.Range.Value2 = oldValues: On Error GoTo 0
    modUtil.ShowError "Assessment", Err.Description
    Resume Done
End Function

Public Function LogMeeting(ByVal assessmentID As Long) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    Dim oldCount As Long
    If Not modIdentity.RequireMutation() Then Exit Function
    Set lo = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    Set lr = modUtil.FindLongRow(lo, "AssessmentID", assessmentID)
    If lr Is Nothing Then Exit Function
    If modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Archived"), False) Then Exit Function
    oldCount = CLng(Val(CStr(modUtil.RowValue(lr, lo, "MeetingsConducted"))))
    modUtil.SetRowValue lr, lo, "MeetingsConducted", oldCount + 1
    modUtil.SetRowValue lr, lo, "ModifiedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, lo, "ModifiedOn", Now
    If Not modAudit.AddAudit("LOG_MEETING", "Assessment", assessmentID, "Count=" & oldCount + 1) Then
        modUtil.SetRowValue lr, lo, "MeetingsConducted", oldCount
        MsgBox "Audit failed; meeting count unchanged.", vbCritical, modUtil.APP_NAME
        Exit Function
    End If
    LogMeeting = True
End Function

Public Function GetAssessmentOwner(ByVal assessmentID As Long, ByRef ownerName As String, ByRef ownerEmail As String, ByRef vendorName As String, ByRef cycle As String) As Boolean
    Dim assessments As ListObject
    Dim vendors As ListObject
    Dim ar As ListRow
    Dim vr As ListRow
    Set assessments = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    Set ar = modUtil.FindLongRow(assessments, "AssessmentID", assessmentID)
    If ar Is Nothing Then Exit Function
    If modUtil.ToBoolean(modUtil.RowValue(ar, assessments, "Archived"), False) Then Exit Function
    Set vendors = modUtil.GetTable(modUtil.TBL_VENDORS)
    Set vr = modUtil.FindLongRow(vendors, "VendorID", CLng(modUtil.RowValue(ar, assessments, "VendorID")))
    If vr Is Nothing Then Exit Function
    ownerName = CStr(modUtil.RowValue(ar, assessments, "BusinessOwnerName"))
    ownerEmail = CStr(modUtil.RowValue(ar, assessments, "BusinessOwnerEmail"))
    vendorName = CStr(modUtil.RowValue(vr, vendors, "VendorName"))
    cycle = CStr(modUtil.RowValue(ar, assessments, "Cycle"))
    GetAssessmentOwner = True
End Function

Public Function ArchiveAssessment(ByVal assessmentID As Long) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    Dim oldValues As Variant
    If Not modIdentity.RequireRole("Manager") Then Exit Function
    Set lo = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    Set lr = modUtil.FindLongRow(lo, "AssessmentID", assessmentID)
    If lr Is Nothing Then Exit Function
    oldValues = lr.Range.Value2
    modUtil.SetRowValue lr, lo, "Archived", True
    modUtil.SetRowValue lr, lo, "ArchivedOn", Now
    modUtil.SetRowValue lr, lo, "ArchivedBy", modIdentity.CurrentUser()
    If Not modAudit.AddAudit("ARCHIVE", "Assessment", assessmentID, "Soft deleted; email events preserved.") Then lr.Range.Value2 = oldValues: MsgBox "Audit failed; assessment unchanged.", vbCritical, modUtil.APP_NAME: Exit Function
    ArchiveAssessment = True
End Function

Public Function HardDeleteAssessment(ByVal assessmentID As Long) As Boolean
    Dim events As ListObject
    Dim assessments As ListObject
    Dim lr As ListRow
    If Not modIdentity.RequireRole("Manager") Then Exit Function
    On Error GoTo BackupFailed
    modBackup.RequireRiskBackup
    Set events = modUtil.GetTable(modUtil.TBL_EMAIL_EVENTS)
    For Each lr In events.ListRows
        If CLng(modUtil.RowValue(lr, events, "AssessmentID")) = assessmentID Then MsgBox "This assessment has email events and can only be archived.", vbExclamation, modUtil.APP_NAME: Exit Function
    Next lr
    Set assessments = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    Set lr = modUtil.FindLongRow(assessments, "AssessmentID", assessmentID)
    If lr Is Nothing Then Exit Function
    If Not modAudit.AddAudit("HARD_DELETE", "Assessment", assessmentID, "No dependent email events.") Then MsgBox "Audit failed; nothing deleted.", vbCritical, modUtil.APP_NAME: Exit Function
    lr.Delete
    HardDeleteAssessment = True
    Exit Function
BackupFailed:
    modUtil.ShowError "Hard delete assessment", Err.Description
End Function

Public Sub RefreshAssessmentView()
    Dim lo As ListObject
    Dim lr As ListRow
    Dim ws As Worksheet
    Dim outRow As Long
    Dim dueDate As Variant
    Set ws = ThisWorkbook.Worksheets("Assessments")
    ws.Range("A5:N1048576").ClearContents
    ws.Range("A4:N4").Value = Array("Assessment ID", "Vendor ID", "Vendor", "Cycle", "Owner", "Owner Email", "Status", "Process Started", "Submission Date", "Completed Date", "Meetings", "Initial Sent", "Last Follow-up", "Next Due")
    outRow = 5
    Set lo = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    For Each lr In lo.ListRows
        If Not modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Archived"), False) Then
            ws.Cells(outRow, 1).Value = modUtil.RowValue(lr, lo, "AssessmentID")
            ws.Cells(outRow, 2).Value = modUtil.RowValue(lr, lo, "VendorID")
            ws.Cells(outRow, 3).Value = VendorNameForID(CLng(modUtil.RowValue(lr, lo, "VendorID")))
            ws.Cells(outRow, 4).Value = modUtil.RowValue(lr, lo, "Cycle")
            ws.Cells(outRow, 5).Value = modUtil.RowValue(lr, lo, "BusinessOwnerName")
            ws.Cells(outRow, 6).Value = modUtil.RowValue(lr, lo, "BusinessOwnerEmail")
            ws.Cells(outRow, 7).Value = modUtil.RowValue(lr, lo, "Status")
            ws.Cells(outRow, 8).Value = (modUtil.StatusRank(CStr(modUtil.RowValue(lr, lo, "Status"))) >= 1)
            ws.Cells(outRow, 9).Value = modUtil.RowValue(lr, lo, "SubmissionDate")
            ws.Cells(outRow, 10).Value = modUtil.RowValue(lr, lo, "CompletedDate")
            ws.Cells(outRow, 11).Value = modUtil.RowValue(lr, lo, "MeetingsConducted")
            ws.Cells(outRow, 12).Value = modEmail.FirstSentDate(CLng(modUtil.RowValue(lr, lo, "AssessmentID")), "Initial")
            ws.Cells(outRow, 13).Value = modEmail.LastSentDate(CLng(modUtil.RowValue(lr, lo, "AssessmentID")), "Followup")
            dueDate = modFollowup.NextFollowupDue(CLng(modUtil.RowValue(lr, lo, "AssessmentID")))
            ws.Cells(outRow, 14).Value = dueDate
            outRow = outRow + 1
        End If
    Next lr
    ws.Range("I5:J" & Application.Max(5, outRow - 1)).NumberFormat = "yyyy-mm-dd"
    ws.Range("L5:N" & Application.Max(5, outRow - 1)).NumberFormat = "yyyy-mm-dd"
End Sub

Public Function VendorNameForID(ByVal vendorID As Long) As String
    Dim lo As ListObject
    Dim lr As ListRow
    Set lo = modUtil.GetTable(modUtil.TBL_VENDORS)
    Set lr = modUtil.FindLongRow(lo, "VendorID", vendorID)
    If Not lr Is Nothing Then VendorNameForID = CStr(modUtil.RowValue(lr, lo, "VendorName"))
End Function
