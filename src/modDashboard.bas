Attribute VB_Name = "modDashboard"
Option Explicit

Public Sub RefreshDashboard()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    Dim counts(0 To 3) As Long
    Dim activeCount As Long
    Dim dueCount As Long
    Dim overdueCount As Long
    Dim outRow As Long
    Dim assessmentID As Long
    Dim dueValue As Variant
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set lo = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    ws.Range("A13:H1048576").ClearContents
    outRow = 13
    For Each lr In lo.ListRows
        If Not modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Archived"), False) Then
            activeCount = activeCount + 1
            If modUtil.StatusRank(CStr(modUtil.RowValue(lr, lo, "Status"))) >= 0 Then counts(modUtil.StatusRank(CStr(modUtil.RowValue(lr, lo, "Status")))) = counts(modUtil.StatusRank(CStr(modUtil.RowValue(lr, lo, "Status")))) + 1
            assessmentID = CLng(modUtil.RowValue(lr, lo, "AssessmentID"))
            dueValue = modFollowup.NextFollowupDue(assessmentID)
            If IsDate(dueValue) And CDate(dueValue) <= Date Then
                dueCount = dueCount + 1
                If CDate(dueValue) < Date Then overdueCount = overdueCount + 1
                ws.Cells(outRow, 1).Value = assessmentID
                ws.Cells(outRow, 2).Value = modAssessments.VendorNameForID(CLng(modUtil.RowValue(lr, lo, "VendorID")))
                ws.Cells(outRow, 3).Value = modUtil.RowValue(lr, lo, "Cycle")
                ws.Cells(outRow, 4).Value = modUtil.RowValue(lr, lo, "BusinessOwnerName")
                ws.Cells(outRow, 5).Value = modUtil.RowValue(lr, lo, "BusinessOwnerEmail")
                ws.Cells(outRow, 6).Value = modUtil.RowValue(lr, lo, "Status")
                ws.Cells(outRow, 7).Value = dueValue
                ws.Cells(outRow, 8).Value = DateDiff("d", CDate(dueValue), Date)
                outRow = outRow + 1
            End If
        End If
    Next lr
    ws.Range("B4").Value = activeCount
    ws.Range("D4").Value = counts(3)
    ws.Range("F4").Value = dueCount
    ws.Range("H4").Value = overdueCount
    ws.Range("J2").Value = "Status": ws.Range("K2").Value = "Count"
    ws.Range("J3").Value = "Not Started": ws.Range("K3").Value = counts(0)
    ws.Range("J4").Value = "In Progress": ws.Range("K4").Value = counts(1)
    ws.Range("J5").Value = "Submitted": ws.Range("K5").Value = counts(2)
    ws.Range("J6").Value = "Completed": ws.Range("K6").Value = counts(3)
    ws.Range("G13:G" & Application.Max(13, outRow - 1)).NumberFormat = "yyyy-mm-dd"
    EnsureStatusChart ws
    modVendors.RefreshVendorView
    modAssessments.RefreshAssessmentView
    ApplyRoleBasedUI
Done:
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then Debug.Print "Dashboard refresh: " & Err.Description
End Sub

Private Sub EnsureStatusChart(ByVal ws As Worksheet)
    Dim chartObject As ChartObject
    On Error Resume Next
    Set chartObject = ws.ChartObjects("chtStatus")
    On Error GoTo 0
    If chartObject Is Nothing Then
        Set chartObject = ws.ChartObjects.Add(ws.Range("J8").Left, ws.Range("J8").Top, 430, 220)
        chartObject.Name = "chtStatus"
    End If
    With chartObject.Chart
        .ChartType = xlColumnClustered
        .SetSourceData ws.Range("J2:K6")
        .HasTitle = True
        .ChartTitle.Text = "Active assessments by status"
        .HasLegend = False
    End With
End Sub

Public Sub ApplyRoleBasedUI()
    Dim ws As Worksheet
    Dim shapeItem As Shape
    Dim managerOnly As Boolean
    managerOnly = (StrComp(modIdentity.CurrentRole(), "Manager", vbTextCompare) = 0 And modIdentity.SessionCanMutate())
    For Each ws In ThisWorkbook.Worksheets
        For Each shapeItem In ws.Shapes
            If Left$(shapeItem.Name, 4) = "mgr_" Then shapeItem.Visible = IIf(managerOnly, msoTrue, msoFalse)
            If Left$(shapeItem.Name, 4) = "mut_" Then shapeItem.Visible = IIf(modIdentity.SessionCanMutate(), msoTrue, msoFalse)
        Next shapeItem
    Next ws
End Sub

Public Sub GoDashboard()
    ThisWorkbook.Worksheets("Dashboard").Activate
End Sub

Public Sub GoVendors()
    ThisWorkbook.Worksheets("Vendors").Activate
End Sub

Public Sub GoAssessments()
    ThisWorkbook.Worksheets("Assessments").Activate
End Sub

Public Sub OpenVendorForm()
    If modIdentity.RequireMutation() Then frmVendor.Show
End Sub

Public Sub OpenAssessmentForm()
    If modIdentity.RequireMutation() Then
        frmAssessment.LoadDefaults
        frmAssessment.Show
    End If
End Sub

Public Sub OpenStatusForm()
    If modIdentity.RequireMutation() Then frmStatus.Show
End Sub

Public Sub OpenSettingsForm()
    If modIdentity.RequireRole("Manager") Then
        frmSettings.LoadSettings
        frmSettings.Show
    End If
End Sub

Public Function SelectedAssessmentID() As Long
    If ActiveSheet.Name = "Assessments" And ActiveCell.Row >= 5 Then SelectedAssessmentID = CLng(Val(CStr(ActiveSheet.Cells(ActiveCell.Row, 1).Value2)))
    If ActiveSheet.Name = "Dashboard" And ActiveCell.Row >= 13 Then SelectedAssessmentID = CLng(Val(CStr(ActiveSheet.Cells(ActiveCell.Row, 1).Value2)))
End Function

Public Sub PrepareSelectedInitial()
    Dim id As Long: id = SelectedAssessmentID()
    If id = 0 Then
        MsgBox "Select an assessment row first.", vbExclamation, modUtil.APP_NAME
    Else
        modEmail.PrepareEmail id, "Initial"
    End If
End Sub

Public Sub PrepareSelectedFollowup()
    Dim id As Long: id = SelectedAssessmentID()
    If id = 0 Then
        MsgBox "Select an assessment row first.", vbExclamation, modUtil.APP_NAME
    Else
        modEmail.PrepareEmail id, "Followup"
    End If
End Sub

Public Sub LogSelectedMeeting()
    Dim id As Long: id = SelectedAssessmentID()
    If id = 0 Then
        MsgBox "Select an assessment row first.", vbExclamation, modUtil.APP_NAME
    Else
        modAssessments.LogMeeting id
    End If
End Sub
