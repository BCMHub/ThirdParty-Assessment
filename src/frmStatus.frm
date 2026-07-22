VERSION 5.00
Begin VB.UserForm frmStatus
   Caption         =   "Assessment Status"
   ClientHeight    =   3600
   ClientWidth     =   6120
   StartUpPosition =   1
   Begin MSForms.TextBox txtAssessmentID
      Height          =   300
      Left            =   2280
      Top             =   360
      Width           =   1800
   End
   Begin MSForms.ComboBox cboStatus
      Height          =   315
      Left            =   2280
      Top             =   840
      Width           =   2760
   End
   Begin MSForms.TextBox txtSubmissionDate
      Height          =   300
      Left            =   2280
      Top             =   1320
      Width           =   1800
   End
   Begin MSForms.TextBox txtCompletedDate
      Height          =   300
      Left            =   2280
      Top             =   1800
      Width           =   1800
   End
   Begin MSForms.CheckBox chkCorrection
      Caption         =   "Manager date correction (same status)"
      Height          =   300
      Left            =   2280
      Top             =   2280
      Width           =   3000
   End
   Begin MSForms.CommandButton cmdApply
      Caption         =   "Apply"
      Height          =   420
      Left            =   3600
      Top             =   2880
      Width           =   1080
   End
   Begin MSForms.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4800
      Top             =   2880
      Width           =   1080
   End
   Begin MSForms.Label lbl1
      Caption         =   "Assessment ID *"
      Left            =   360
      Top             =   420
      Width           =   1680
   End
   Begin MSForms.Label lbl2
      Caption         =   "Target status *"
      Left            =   360
      Top             =   900
      Width           =   1680
   End
   Begin MSForms.Label lbl3
      Caption         =   "Submission (yyyy-mm-dd)"
      Left            =   360
      Top             =   1380
      Width           =   1800
   End
   Begin MSForms.Label lbl4
      Caption         =   "Completed (yyyy-mm-dd)"
      Left            =   360
      Top             =   1860
      Width           =   1800
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSubmitting As Boolean

Private Sub UserForm_Initialize()
    cboStatus.Clear
    cboStatus.AddItem "Not Started"
    cboStatus.AddItem "In Progress"
    cboStatus.AddItem "Submitted"
    cboStatus.AddItem "Completed"
    If modDashboard.SelectedAssessmentID() > 0 Then
        txtAssessmentID.Value = modDashboard.SelectedAssessmentID()
        LoadCurrentAssessment CLng(txtAssessmentID.Value)
    End If
    chkCorrection.Visible = (StrComp(modIdentity.CurrentRole(), "Manager", vbTextCompare) = 0)
End Sub

Private Sub txtAssessmentID_AfterUpdate()
    If IsNumeric(txtAssessmentID.Value) Then LoadCurrentAssessment CLng(txtAssessmentID.Value)
End Sub

Private Sub LoadCurrentAssessment(ByVal assessmentID As Long)
    Dim lo As ListObject
    Dim lr As ListRow
    Set lo = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    Set lr = modUtil.FindLongRow(lo, "AssessmentID", assessmentID)
    If lr Is Nothing Then Exit Sub
    cboStatus.Value = CStr(modUtil.RowValue(lr, lo, "Status"))
    If IsDate(modUtil.RowValue(lr, lo, "SubmissionDate")) Then txtSubmissionDate.Value = Format$(CDate(modUtil.RowValue(lr, lo, "SubmissionDate")), "yyyy-mm-dd") Else txtSubmissionDate.Value = vbNullString
    If IsDate(modUtil.RowValue(lr, lo, "CompletedDate")) Then txtCompletedDate.Value = Format$(CDate(modUtil.RowValue(lr, lo, "CompletedDate")), "yyyy-mm-dd") Else txtCompletedDate.Value = vbNullString
End Sub

Private Sub cmdApply_Click()
    Dim submissionValue As Variant, completedValue As Variant, mode As String
    If mSubmitting Or Not IsNumeric(txtAssessmentID.Value) Or cboStatus.ListIndex < 0 Then Exit Sub
    If Len(Trim$(txtSubmissionDate.Value)) = 0 Then
        submissionValue = Empty
    ElseIf Not IsDate(txtSubmissionDate.Value) Then
        MsgBox "Invalid submission date.", vbExclamation
        Exit Sub
    Else
        submissionValue = CDate(txtSubmissionDate.Value)
    End If
    If Len(Trim$(txtCompletedDate.Value)) = 0 Then
        completedValue = Empty
    ElseIf Not IsDate(txtCompletedDate.Value) Then
        MsgBox "Invalid completed date.", vbExclamation
        Exit Sub
    Else
        completedValue = CDate(txtCompletedDate.Value)
    End If
    mode = IIf(chkCorrection.Value, modStatus.MODE_CORRECTION, modStatus.MODE_TRANSITION)
    mSubmitting = True: cmdApply.Enabled = False
    If modStatus.ApplyTransition(CLng(txtAssessmentID.Value), cboStatus.Value, mode, submissionValue, completedValue) Then Unload Me
    cmdApply.Enabled = True: mSubmitting = False
End Sub

Private Sub cmdCancel_Click(): Unload Me: End Sub
