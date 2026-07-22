VERSION 5.00
Begin VB.UserForm frmBatchPreview
   Caption         =   "Batch Email Preview"
   ClientHeight    =   5640
   ClientWidth     =   8400
   StartUpPosition =   1
   Begin MSForms.Label lblSummary
      Caption         =   "Summary"
      Height          =   720
      Left            =   360
      Top             =   240
      Width           =   7680
      WordWrap        =   -1
   End
   Begin MSForms.ListBox lstItems
      Height          =   3480
      Left            =   360
      Top             =   1080
      Width           =   7680
   End
   Begin MSForms.CommandButton cmdConfirm
      Caption         =   "Confirm"
      Height          =   480
      Left            =   5760
      Top             =   4800
      Width           =   1080
   End
   Begin MSForms.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   480
      Left            =   6960
      Top             =   4800
      Width           =   1080
   End
End
Attribute VB_Name = "frmBatchPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Confirmed As Boolean

Public Sub LoadPreview(ByRef labels() As String, ByVal count As Long, ByVal summary As String)
    Dim i As Long
    Confirmed = False
    lblSummary.Caption = summary
    lstItems.Clear
    For i = 1 To count: lstItems.AddItem labels(i): Next i
End Sub

Private Sub cmdConfirm_Click(): Confirmed = True: Me.Hide: End Sub
Private Sub cmdCancel_Click(): Confirmed = False: Me.Hide: End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Confirmed = False
End Sub
