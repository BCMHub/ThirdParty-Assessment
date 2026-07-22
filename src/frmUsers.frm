VERSION 5.00
Begin VB.UserForm frmUsers
   Caption         =   "Manage Users"
   ClientHeight    =   5040
   ClientWidth     =   7800
   StartUpPosition =   1
   Begin MSForms.ListBox lstUsers
      Height          =   2040
      Left            =   360
      Top             =   360
      Width           =   7080
   End
   Begin MSForms.TextBox txtIdentity
      Height          =   300
      Left            =   2160
      Top             =   2760
      Width           =   3000
   End
   Begin MSForms.TextBox txtDisplayName
      Height          =   300
      Left            =   2160
      Top             =   3240
      Width           =   3000
   End
   Begin MSForms.ComboBox cboRole
      Height          =   315
      Left            =   2160
      Top             =   3720
      Width           =   1800
   End
   Begin MSForms.CheckBox chkActive
      Caption         =   "Active"
      Height          =   300
      Left            =   4200
      Top             =   3720
      Width           =   960
   End
   Begin MSForms.CommandButton cmdSave
      Caption         =   "Add / Update"
      Height          =   420
      Left            =   5520
      Top             =   3720
      Width           =   1440
   End
   Begin MSForms.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   420
      Left            =   6360
      Top             =   4320
      Width           =   1080
   End
   Begin MSForms.Label lbl1
      Caption         =   "DOMAIN\\USERNAME"
      Left            =   360
      Top             =   2820
      Width           =   1680
   End
   Begin MSForms.Label lbl2
      Caption         =   "Display name"
      Left            =   360
      Top             =   3300
      Width           =   1680
   End
   Begin MSForms.Label lbl3
      Caption         =   "Role"
      Left            =   360
      Top             =   3780
      Width           =   1680
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSubmitting As Boolean

Public Sub LoadUsers()
    Dim lo As ListObject, lr As ListRow
    lstUsers.Clear: lstUsers.ColumnCount = 4: lstUsers.ColumnWidths = "170 pt;150 pt;70 pt;45 pt"
    cboRole.Clear
    cboRole.AddItem "Assessor"
    cboRole.AddItem "Manager"
    chkActive.Value = True
    Set lo = modUtil.GetTable(modUtil.TBL_USERS)
    For Each lr In lo.ListRows
        lstUsers.AddItem CStr(modUtil.RowValue(lr, lo, "WindowsUsername"))
        lstUsers.List(lstUsers.ListCount - 1, 1) = CStr(modUtil.RowValue(lr, lo, "DisplayName"))
        lstUsers.List(lstUsers.ListCount - 1, 2) = CStr(modUtil.RowValue(lr, lo, "Role"))
        lstUsers.List(lstUsers.ListCount - 1, 3) = CStr(modUtil.RowValue(lr, lo, "Active"))
    Next lr
End Sub

Private Sub lstUsers_Click()
    If lstUsers.ListIndex < 0 Then Exit Sub
    txtIdentity.Value = lstUsers.List(lstUsers.ListIndex, 0)
    txtDisplayName.Value = lstUsers.List(lstUsers.ListIndex, 1)
    cboRole.Value = lstUsers.List(lstUsers.ListIndex, 2)
    chkActive.Value = modUtil.ToBoolean(lstUsers.List(lstUsers.ListIndex, 3), False)
End Sub

Private Sub cmdSave_Click()
    Dim lo As ListObject, existing As ListRow, ok As Boolean
    If mSubmitting Then Exit Sub
    mSubmitting = True: cmdSave.Enabled = False
    Set lo = modUtil.GetTable(modUtil.TBL_USERS)
    Set existing = modUtil.FindTextRow(lo, "WindowsUsername", txtIdentity.Value)
    If existing Is Nothing Then ok = modUsers.AddUser(txtIdentity.Value, txtDisplayName.Value, cboRole.Value) Else ok = modUsers.UpdateUser(txtIdentity.Value, txtDisplayName.Value, cboRole.Value, chkActive.Value)
    cmdSave.Enabled = True: mSubmitting = False
    If ok Then LoadUsers
End Sub

Private Sub cmdClose_Click(): Unload Me: End Sub
