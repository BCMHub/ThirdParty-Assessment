Attribute VB_Name = "modStatus"
Option Explicit

Public Const MODE_TRANSITION As String = "Transition"
Public Const MODE_CORRECTION As String = "Correction"
Private mBusy As Boolean

Public Function ApplyTransition(ByVal assessmentID As Long, ByVal newStatus As String, Optional ByVal mode As String = MODE_TRANSITION, _
    Optional ByVal newSubmissionDate As Variant, Optional ByVal newCompletedDate As Variant) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    Dim oldStatus As String
    Dim oldSubmission As Variant
    Dim oldCompleted As Variant
    Dim targetSubmission As Variant
    Dim targetCompleted As Variant
    Dim oldRow As Variant
    Dim isCorrection As Boolean
    Dim detail As String
    If mBusy Then Exit Function
    If Not modIdentity.RequireMutation() Then Exit Function
    newStatus = CanonicalStatus(newStatus)
    If modUtil.StatusRank(newStatus) < 0 Then MsgBox "Invalid status.", vbExclamation, modUtil.APP_NAME: Exit Function
    Set lo = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    Set lr = modUtil.FindLongRow(lo, "AssessmentID", assessmentID)
    If lr Is Nothing Then MsgBox "Assessment not found.", vbExclamation, modUtil.APP_NAME: Exit Function
    If modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Archived"), False) Then MsgBox "Archived assessments cannot be changed.", vbExclamation, modUtil.APP_NAME: Exit Function
    oldStatus = CStr(modUtil.RowValue(lr, lo, "Status"))
    oldSubmission = modUtil.RowValue(lr, lo, "SubmissionDate")
    oldCompleted = modUtil.RowValue(lr, lo, "CompletedDate")
    isCorrection = (StrComp(mode, MODE_CORRECTION, vbTextCompare) = 0)
    If modUtil.StatusRank(newStatus) < modUtil.StatusRank(oldStatus) Or isCorrection Then
        If Not modIdentity.RequireRole("Manager") Then Exit Function
    End If
    If StrComp(newStatus, oldStatus, vbTextCompare) = 0 And Not isCorrection Then
        ApplyTransition = True
        Exit Function
    End If
    If StrComp(newStatus, oldStatus, vbTextCompare) <> 0 And isCorrection Then
        MsgBox "Correction mode is only for date changes while status stays the same.", vbExclamation, modUtil.APP_NAME
        Exit Function
    End If
    If IsMissing(newSubmissionDate) Then targetSubmission = oldSubmission Else targetSubmission = newSubmissionDate
    If IsMissing(newCompletedDate) Then targetCompleted = oldCompleted Else targetCompleted = newCompletedDate
    If Not ResolveTargetDates(newStatus, oldStatus, isCorrection, targetSubmission, targetCompleted) Then Exit Function
    If StrComp(newStatus, oldStatus, vbTextCompare) = 0 And modUtil.ValuesEqual(oldSubmission, targetSubmission) And modUtil.ValuesEqual(oldCompleted, targetCompleted) Then
        ApplyTransition = True
        Exit Function
    End If
    If Not ValidateInvariant(newStatus, targetSubmission, targetCompleted, detail) Then
        MsgBox detail, vbExclamation, modUtil.APP_NAME
        Exit Function
    End If
    oldRow = lr.Range.Value2
    On Error GoTo Failed
    mBusy = True
    modUtil.SetRowValue lr, lo, "Status", newStatus
    modUtil.SetRowValue lr, lo, "SubmissionDate", targetSubmission
    modUtil.SetRowValue lr, lo, "CompletedDate", targetCompleted
    modUtil.SetRowValue lr, lo, "ModifiedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, lo, "ModifiedOn", Now
    detail = "OldStatus=" & oldStatus & "; NewStatus=" & newStatus & "; SubmissionDate=" & DateText(targetSubmission) & "; CompletedDate=" & DateText(targetCompleted)
    If Not modAudit.AddAudit(IIf(isCorrection, "STATUS_CORRECTION", "STATUS_TRANSITION"), "Assessment", assessmentID, detail) Then
        lr.Range.Value2 = oldRow
        Err.Raise vbObjectError + 1600, "modStatus.ApplyTransition", "The audit write failed; the assessment row was rolled back."
    End If
    ApplyTransition = True
    modDashboard.RefreshDashboard
Done:
    mBusy = False
    Exit Function
Failed:
    On Error Resume Next: lr.Range.Value2 = oldRow: On Error GoTo 0
    modUtil.ShowError "Status", Err.Description
    Resume Done
End Function

Private Function ResolveTargetDates(ByVal targetStatus As String, ByVal oldStatus As String, ByVal isCorrection As Boolean, ByRef submissionValue As Variant, ByRef completedValue As Variant) As Boolean
    Select Case targetStatus
        Case "Not Started", "In Progress"
            submissionValue = Empty
            completedValue = Empty
        Case "Submitted"
            completedValue = Empty
            If Not IsDate(submissionValue) Then
                If isCorrection Then submissionValue = PromptForDate("Submission date", Date) Else submissionValue = PromptForDate("Submission date", Date)
                If IsEmpty(submissionValue) Then Exit Function
            End If
        Case "Completed"
            If Not IsDate(submissionValue) Then
                submissionValue = PromptForDate("Submission date", Date)
                If IsEmpty(submissionValue) Then Exit Function
            End If
            If Not IsDate(completedValue) Then
                completedValue = PromptForDate("Completed date", Date)
                If IsEmpty(completedValue) Then Exit Function
            End If
    End Select
    ResolveTargetDates = True
End Function

Private Function PromptForDate(ByVal labelText As String, ByVal defaultDate As Date) As Variant
    Dim answer As Variant
    answer = Application.InputBox(labelText & " (yyyy-mm-dd)", modUtil.APP_NAME, Format$(defaultDate, "yyyy-mm-dd"), Type:=2)
    If VarType(answer) = vbBoolean And answer = False Then
        PromptForDate = Empty
    ElseIf IsDate(answer) Then
        PromptForDate = CDate(answer)
    Else
        MsgBox labelText & " is not a valid date.", vbExclamation, modUtil.APP_NAME
        PromptForDate = Empty
    End If
End Function

Public Function ValidateInvariant(ByVal statusName As String, ByVal submissionValue As Variant, ByVal completedValue As Variant, ByRef reason As String) As Boolean
    Select Case CanonicalStatus(statusName)
        Case "Not Started", "In Progress"
            If IsDate(submissionValue) Or IsDate(completedValue) Then reason = "Not Started/In Progress require both dates to be blank.": Exit Function
        Case "Submitted"
            If Not IsDate(submissionValue) Or IsDate(completedValue) Then reason = "Submitted requires a submission date and a blank completed date.": Exit Function
            If CDate(submissionValue) > Date Then reason = "Submission date cannot be in the future.": Exit Function
        Case "Completed"
            If Not IsDate(submissionValue) Or Not IsDate(completedValue) Then reason = "Completed requires both dates.": Exit Function
            If CDate(submissionValue) > CDate(completedValue) Then reason = "Submission date must be on or before completed date.": Exit Function
            If CDate(completedValue) > Date Then reason = "Completed date cannot be in the future.": Exit Function
        Case Else
            reason = "Unknown status.": Exit Function
    End Select
    ValidateInvariant = True
End Function

Public Function CanonicalStatus(ByVal statusName As String) As String
    Select Case LCase$(Trim$(statusName))
        Case "not started": CanonicalStatus = "Not Started"
        Case "in progress": CanonicalStatus = "In Progress"
        Case "submitted": CanonicalStatus = "Submitted"
        Case "completed": CanonicalStatus = "Completed"
    End Select
End Function

Private Function DateText(ByVal value As Variant) As String
    If IsDate(value) Then DateText = Format$(CDate(value), "yyyy-mm-dd") Else DateText = "<blank>"
End Function
