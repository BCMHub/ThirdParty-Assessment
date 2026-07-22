VERSION 5.00
Begin VB.UserForm frmSettings
   Caption         =   "Settings"
   ClientHeight    =   6840
   ClientWidth     =   8280
   StartUpPosition =   1
   Begin MSForms.ListBox lstSettings
      Height          =   4440
      Left            =   360
      Top             =   360
      Width           =   7560
   End
   Begin MSForms.TextBox txtName
      Height          =   300
      Left            =   1800
      Top             =   5160
      Width           =   2400
   End
   Begin MSForms.TextBox txtValue
      Height          =   600
      Left            =   1800
      MultiLine       =   -1
      Top             =   5640
      Width           =   4440
   End
   Begin MSForms.CommandButton cmdSave
      Caption         =   "Save setting"
      Height          =   420
      Left            =   6480
      Top             =   5640
      Width           =   1320
   End
   Begin MSForms.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   420
      Left            =   6480
      Top             =   6120
      Width           =   1320
   End
   Begin MSForms.Label lbl1
      Caption         =   "Setting"
      Left            =   360
      Top             =   5220
      Width           =   1200
   End
   Begin MSForms.Label lbl2
      Caption         =   "Value"
      Left            =   360
      Top             =   5700
      Width           =   1200
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSubmitting As Boolean

Public Sub LoadSettings()
    Dim lo As ListObject, lr As ListRow
    lstSettings.Clear: lstSettings.ColumnCount = 3: lstSettings.ColumnWidths = "150 pt;250 pt;300 pt"
    Set lo = modUtil.GetTable(modUtil.TBL_SETTINGS)
    For Each lr In lo.ListRows
        lstSettings.AddItem CStr(modUtil.RowValue(lr, lo, "SettingName"))
        lstSettings.List(lstSettings.ListCount - 1, 1) = CStr(modUtil.RowValue(lr, lo, "SettingValue"))
        lstSettings.List(lstSettings.ListCount - 1, 2) = CStr(modUtil.RowValue(lr, lo, "Description"))
    Next lr
End Sub

Private Sub lstSettings_Click()
    If lstSettings.ListIndex < 0 Then Exit Sub
    txtName.Value = lstSettings.List(lstSettings.ListIndex, 0)
    txtValue.Value = lstSettings.List(lstSettings.ListIndex, 1)
    txtName.Enabled = False
End Sub

Private Sub cmdSave_Click()
    Dim ok As Boolean
    If mSubmitting Then Exit Sub
    mSubmitting = True
    cmdSave.Enabled = False
    ok = modSettings.UpdateSetting(txtName.Value, txtValue.Value)
    cmdSave.Enabled = True
    mSubmitting = False
    If ok Then LoadSettings
End Sub

Private Sub cmdClose_Click(): Unload Me: End Sub
