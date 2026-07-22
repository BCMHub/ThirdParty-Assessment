VERSION 5.00
Begin VB.UserForm frmAssessment
   Caption         =   "Assessment"
   ClientHeight    =   5040
   ClientWidth     =   7200
   StartUpPosition =   1
   Begin MSForms.TextBox txtAssessmentID
      Height          =   300
      Left            =   1920
      Top             =   240
      Width           =   1680
   End
   Begin MSForms.ComboBox cboVendor
      Height          =   315
      Left            =   1920
      Top             =   720
      Width           =   4680
   End
   Begin MSForms.TextBox txtCycle
      Height          =   300
      Left            =   1920
      Top             =   1200
      Width           =   2400
   End
   Begin MSForms.TextBox txtOwnerName
      Height          =   300
      Left            =   1920
      Top             =   1680
      Width           =   4680
   End
   Begin MSForms.TextBox txtOwnerEmail
      Height          =   300
      Left            =   1920
      Top             =   2160
      Width           =   4680
   End
   Begin MSForms.TextBox txtNotes
      Height          =   960
      Left            =   1920
      MultiLine       =   -1
      Top             =   2640
      Width           =   4680
   End
   Begin MSForms.CommandButton cmdSave
      Caption         =   "Save"
      Height          =   420
      Left            =   4320
      Top             =   4080
      Width           =   1080
   End
   Begin MSForms.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   420
      Left            =   5520
      Top             =   4080
      Width           =   1080
   End
   Begin MSForms.Label lbl1
      Caption         =   "Assessment ID"
      Left            =   240
      Top             =   300
      Width           =   1440
   End
   Begin MSForms.Label lbl2
      Caption         =   "Vendor *"
      Left            =   240
      Top             =   780
      Width           =   1440
   End
   Begin MSForms.Label lbl3
      Caption         =   "Cycle *"
      Left            =   240
      Top             =   1260
      Width           =   1440
   End
   Begin MSForms.Label lbl4
      Caption         =   "Owner snapshot *"
      Left            =   240
      Top             =   1740
      Width           =   1440
   End
   Begin MSForms.Label lbl5
      Caption         =   "Owner email *"
      Left            =   240
      Top             =   2220
      Width           =   1440
   End
   Begin MSForms.Label lbl6
      Caption         =   "Notes"
      Left            =   240
      Top             =   2700
      Width           =   1440
   End
End
Attribute VB_Name = "frmAssessment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSubmitting As Boolean

Public Sub LoadDefaults()
    Dim lo As ListObject, lr As ListRow
    cboVendor.Clear
    cboVendor.ColumnCount = 2
    cboVendor.ColumnWidths = "60 pt;260 pt"
    Set lo = modUtil.GetTable(modUtil.TBL_VENDORS)
    For Each lr In lo.ListRows
        If Not modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Archived"), False) Then
            cboVendor.AddItem CStr(modUtil.RowValue(lr, lo, "VendorID"))
            cboVendor.List(cboVendor.ListCount - 1, 1) = CStr(modUtil.RowValue(lr, lo, "VendorName"))
        End If
    Next lr
    txtAssessmentID.Enabled = False
End Sub

Private Sub cboVendor_Change()
    Dim lo As ListObject, lr As ListRow
    If cboVendor.ListIndex < 0 Then Exit Sub
    Set lo = modUtil.GetTable(modUtil.TBL_VENDORS)
    Set lr = modUtil.FindLongRow(lo, "VendorID", CLng(cboVendor.List(cboVendor.ListIndex, 0)))
    If Not lr Is Nothing Then
        txtOwnerName.Value = modUtil.RowValue(lr, lo, "BusinessOwnerName")
        txtOwnerEmail.Value = modUtil.RowValue(lr, lo, "BusinessOwnerEmail")
    End If
End Sub

Public Sub LoadAssessment(ByVal assessmentID As Long)
    Dim lo As ListObject, lr As ListRow, i As Long
    LoadDefaults
    Set lo = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    Set lr = modUtil.FindLongRow(lo, "AssessmentID", assessmentID)
    If lr Is Nothing Then Exit Sub
    txtAssessmentID.Value = assessmentID
    For i = 0 To cboVendor.ListCount - 1
        If CLng(cboVendor.List(i, 0)) = CLng(modUtil.RowValue(lr, lo, "VendorID")) Then cboVendor.ListIndex = i: Exit For
    Next i
    cboVendor.Enabled = False
    txtCycle.Value = modUtil.RowValue(lr, lo, "Cycle")
    txtOwnerName.Value = modUtil.RowValue(lr, lo, "BusinessOwnerName")
    txtOwnerEmail.Value = modUtil.RowValue(lr, lo, "BusinessOwnerEmail")
    txtNotes.Value = modUtil.RowValue(lr, lo, "Notes")
End Sub

Private Sub cmdSave_Click()
    Dim ok As Boolean, vendorID As Long
    If mSubmitting Or cboVendor.ListIndex < 0 Then Exit Sub
    mSubmitting = True: cmdSave.Enabled = False
    vendorID = CLng(cboVendor.List(cboVendor.ListIndex, 0))
    If Len(Trim$(txtAssessmentID.Value)) = 0 Then
        ok = (modAssessments.CreateAssessment(vendorID, txtCycle.Value, txtNotes.Value, txtOwnerName.Value, txtOwnerEmail.Value) > 0)
    Else
        ok = modAssessments.UpdateAssessmentDetails(CLng(txtAssessmentID.Value), txtCycle.Value, txtOwnerName.Value, txtOwnerEmail.Value, txtNotes.Value)
    End If
    cmdSave.Enabled = True: mSubmitting = False
    If ok Then Unload Me
End Sub

Private Sub cmdCancel_Click(): Unload Me: End Sub
