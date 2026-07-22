Attribute VB_Name = "modUsers"
Option Explicit

Private mBusy As Boolean

Public Function AddUser(ByVal windowsUsername As String, ByVal displayName As String, ByVal roleName As String) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    If mBusy Then Exit Function
    If Not modIdentity.RequireRole("Manager") Then Exit Function
    windowsUsername = UCase$(Trim$(windowsUsername))
    roleName = CanonicalRole(roleName)
    If InStr(windowsUsername, "\") <= 1 Then MsgBox "Use DOMAIN\USERNAME.", vbExclamation, modUtil.APP_NAME: Exit Function
    If Len(roleName) = 0 Then MsgBox "Role must be Assessor or Manager.", vbExclamation, modUtil.APP_NAME: Exit Function
    Set lo = modUtil.GetTable(modUtil.TBL_USERS)
    If Not modUtil.FindTextRow(lo, "WindowsUsername", windowsUsername) Is Nothing Then MsgBox "That Windows identity already exists.", vbExclamation, modUtil.APP_NAME: Exit Function
    On Error GoTo Failed
    mBusy = True
    Set lr = lo.ListRows.Add
    modUtil.SetRowValue lr, lo, "WindowsUsername", windowsUsername
    modUtil.SetRowValue lr, lo, "DisplayName", modUtil.NeutralizeImportedText(Trim$(displayName))
    modUtil.SetRowValue lr, lo, "Role", roleName
    modUtil.SetRowValue lr, lo, "Active", True
    modUtil.SetRowValue lr, lo, "CreatedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, lo, "CreatedOn", Now
    If Not modAudit.AddAudit("CREATE", "User", windowsUsername, "Role=" & roleName) Then lr.Delete: Err.Raise vbObjectError + 1900, , "Audit failed; user not added."
    AddUser = True
    modIdentity.RefreshAuthorization
Done:
    mBusy = False
    Exit Function
Failed:
    On Error Resume Next: If Not lr Is Nothing Then lr.Delete
    On Error GoTo 0
    modUtil.ShowError "Users", Err.Description
    Resume Done
End Function

Public Function UpdateUser(ByVal windowsUsername As String, ByVal displayName As String, ByVal roleName As String, ByVal active As Boolean) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    Dim oldRow As Variant
    If mBusy Then Exit Function
    If Not modIdentity.RequireRole("Manager") Then Exit Function
    windowsUsername = UCase$(Trim$(windowsUsername))
    roleName = CanonicalRole(roleName)
    If Len(roleName) = 0 Then Exit Function
    Set lo = modUtil.GetTable(modUtil.TBL_USERS)
    Set lr = modUtil.FindTextRow(lo, "WindowsUsername", windowsUsername)
    If lr Is Nothing Then MsgBox "User not found.", vbExclamation, modUtil.APP_NAME: Exit Function
    oldRow = lr.Range.Value2
    On Error GoTo Failed
    mBusy = True
    modUtil.SetRowValue lr, lo, "DisplayName", modUtil.NeutralizeImportedText(Trim$(displayName))
    modUtil.SetRowValue lr, lo, "Role", roleName
    modUtil.SetRowValue lr, lo, "Active", active
    If CountActiveManagers(lo) < 1 Then lr.Range.Value2 = oldRow: MsgBox "At least one active Manager is required.", vbCritical, modUtil.APP_NAME: GoTo Done
    If Not modAudit.AddAudit("UPDATE", "User", windowsUsername, "Role=" & roleName & "; Active=" & CStr(active)) Then lr.Range.Value2 = oldRow: Err.Raise vbObjectError + 1901, , "Audit failed; user unchanged."
    UpdateUser = True
    modIdentity.RefreshAuthorization
Done:
    mBusy = False
    Exit Function
Failed:
    On Error Resume Next: lr.Range.Value2 = oldRow: On Error GoTo 0
    modUtil.ShowError "Users", Err.Description
    Resume Done
End Function

Private Function CountActiveManagers(ByVal lo As ListObject) As Long
    Dim lr As ListRow
    For Each lr In lo.ListRows
        If modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Active"), False) And StrComp(CStr(modUtil.RowValue(lr, lo, "Role")), "Manager", vbTextCompare) = 0 Then CountActiveManagers = CountActiveManagers + 1
    Next lr
End Function

Private Function CanonicalRole(ByVal roleName As String) As String
    If StrComp(roleName, "Assessor", vbTextCompare) = 0 Then CanonicalRole = "Assessor"
    If StrComp(roleName, "Manager", vbTextCompare) = 0 Then CanonicalRole = "Manager"
End Function

Public Sub ShowUsersForm()
    If Not modIdentity.RequireRole("Manager") Then Exit Sub
    frmUsers.LoadUsers
    frmUsers.Show
End Sub
