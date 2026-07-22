VERSION 5.00
Begin VB.UserForm frmImportReview
   Caption         =   "Import Review"
   ClientHeight    =   4200
   ClientWidth     =   6480
   StartUpPosition =   1
   Begin MSForms.Label lblSummary
      Caption         =   "Review"
      Height          =   1320
      Left            =   360
      Top             =   360
      Width           =   5760
      WordWrap        =   -1
   End
   Begin MSForms.CheckBox chkOwnerName
      Caption         =   "Update owner name"
      Height          =   300
      Left            =   360
      Top             =   1680
      Width           =   1800
   End
   Begin MSForms.CheckBox chkContactPerson
      Caption         =   "Update contact person"
      Height          =   300
      Left            =   2280
      Top             =   1680
      Width           =   1920
   End
   Begin MSForms.CheckBox chkContactEmail
      Caption         =   "Update contact email"
      Height          =   300
      Left            =   4320
      Top             =   1680
      Width           =   1800
   End
   Begin MSForms.CheckBox chkContactPhone
      Caption         =   "Update contact phone"
      Height          =   300
      Left            =   360
      Top             =   2160
      Width           =   1800
   End
   Begin MSForms.CheckBox chkNotes
      Caption         =   "Update notes"
      Height          =   300
      Left            =   2280
      Top             =   2160
      Width           =   1800
   End
   Begin MSForms.CommandButton cmdCommit
      Caption         =   "Commit valid rows"
      Height          =   480
      Left            =   3360
      Top             =   3120
      Width           =   1560
   End
   Begin MSForms.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   480
      Left            =   5040
      Top             =   3120
      Width           =   1080
   End
End
Attribute VB_Name = "frmImportReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSubmitting As Boolean

Public Sub LoadReview(ByVal validCount As Long, ByVal errorCount As Long, ByVal sourceFile As String, ByVal importMode As String)
    lblSummary.Caption = "Source: " & sourceFile & vbCrLf & "Mode: " & importMode & vbCrLf & "Valid: " & validCount & " | Errors: " & errorCount & vbCrLf & "Inspect ImportStaging. Only valid rows are eligible."
    cmdCommit.Enabled = (validCount > 0)
    Call SetUpdateChoicesVisible(importMode = "UPDATE")
End Sub

Private Sub SetUpdateChoicesVisible(ByVal isUpdate As Boolean)
    chkOwnerName.Visible = isUpdate
    chkContactPerson.Visible = isUpdate
    chkContactEmail.Visible = isUpdate
    chkContactPhone.Visible = isUpdate
    chkNotes.Visible = isUpdate
    chkOwnerName.Value = isUpdate
    chkContactPerson.Value = isUpdate
    chkContactEmail.Value = isUpdate
    chkContactPhone.Value = isUpdate
    chkNotes.Value = isUpdate
End Sub

Public Function UpdateOwnerName() As Boolean: UpdateOwnerName = chkOwnerName.Value: End Function
Public Function UpdateContactPerson() As Boolean: UpdateContactPerson = chkContactPerson.Value: End Function
Public Function UpdateContactEmail() As Boolean: UpdateContactEmail = chkContactEmail.Value: End Function
Public Function UpdateContactPhone() As Boolean: UpdateContactPhone = chkContactPhone.Value: End Function
Public Function UpdateNotes() As Boolean: UpdateNotes = chkNotes.Value: End Function

Private Sub cmdCommit_Click()
    If mSubmitting Then Exit Sub
    mSubmitting = True: cmdCommit.Enabled = False
    Me.Hide
    modImport.CommitStagedImport
    Unload Me
End Sub

Private Sub cmdCancel_Click(): Unload Me: End Sub
