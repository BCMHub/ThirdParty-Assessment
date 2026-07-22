VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo Failed
    Application.ScreenUpdating = False
    modIdentity.InitializeSession
    If Not ThisWorkbook.ReadOnly And modIdentity.SessionCanMutate() Then
        If Not modBackup.EnsureDailyBackup(False, True) Then Debug.Print "Daily backup failed; high-risk actions remain separately blocked by forced backup."
    End If
    If modEnv.Check(False) Then
        If modIdentity.SessionCanMutate() Then modEmail.Reconcile False
    Else
        MsgBox "Startup self-test failed. The workbook is in Browse mode." & vbCrLf & modIdentity.MutationBlockReason(), vbCritical, modUtil.APP_NAME
    End If
    modDashboard.RefreshDashboard
    ThisWorkbook.Worksheets("Dashboard").Activate
Done:
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    modIdentity.BlockMutations "Startup error: " & Err.Description
    MsgBox "Startup failed closed; mutations are blocked." & vbCrLf & Err.Description, vbCritical, modUtil.APP_NAME
    Resume Done
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    modDashboard.RefreshDashboard
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.StatusBar = False
End Sub
