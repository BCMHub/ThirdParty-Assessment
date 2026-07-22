Attribute VB_Name = "modIdentity"
Option Explicit

Private mSessionIdentity As String
Private mSessionRole As String
Private mMutationBlocked As Boolean
Private mBlockReason As String

Public Sub InitializeSession()
    mSessionIdentity = ResolveWindowsIdentity()
    RefreshAuthorization
End Sub

Public Function CurrentUser() As String
    If Len(mSessionIdentity) = 0 Then mSessionIdentity = ResolveWindowsIdentity()
    CurrentUser = mSessionIdentity
End Function

Public Function ResolveWindowsIdentity() As String
    Dim domainName As String
    Dim userName As String
    Dim network As Object
    domainName = Trim$(Environ$("USERDOMAIN"))
    userName = Trim$(Environ$("USERNAME"))
    If Len(domainName) = 0 Or Len(userName) = 0 Then
        On Error Resume Next
        Set network = CreateObject("WScript.Network")
        If Not network Is Nothing Then
            domainName = Trim$(CStr(network.UserDomain))
            userName = Trim$(CStr(network.UserName))
        End If
        On Error GoTo 0
    End If
    If Len(domainName) > 0 And Len(userName) > 0 Then ResolveWindowsIdentity = UCase$(domainName & "\" & userName)
End Function

Public Sub RefreshAuthorization()
    Dim lo As ListObject
    Dim lr As ListRow
    Dim activeCount As Long
    Dim managerCount As Long
    Dim identityValue As String
    Dim roleValue As String
    Dim activeValue As Boolean
    Dim seen As Object
    mSessionRole = vbNullString
    mMutationBlocked = False
    mBlockReason = vbNullString
    If Len(CurrentUser()) = 0 Then
        BlockMutations "Windows identity could not be resolved."
        Exit Sub
    End If
    If StrComp(modSettings.GetSetting("BootstrapState", vbNullString), "Bootstrapped", vbBinaryCompare) <> 0 Then
        BlockMutations "Bootstrap state is missing, pending, unknown, or corrupt."
        Exit Sub
    End If
    On Error GoTo Corrupt
    Set lo = modUtil.GetTable(modUtil.TBL_USERS)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    For Each lr In lo.ListRows
        identityValue = UCase$(Trim$(CStr(modUtil.RowValue(lr, lo, "WindowsUsername"))))
        roleValue = Trim$(CStr(modUtil.RowValue(lr, lo, "Role")))
        If Not IsRecognizedBoolean(modUtil.RowValue(lr, lo, "Active")) Then GoTo Corrupt
        activeValue = modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Active"), False)
        If InStr(identityValue, "\") <= 1 Then GoTo Corrupt
        If seen.Exists(identityValue) Then GoTo Corrupt Else seen(identityValue) = True
        If StrComp(roleValue, "Assessor", vbTextCompare) <> 0 And StrComp(roleValue, "Manager", vbTextCompare) <> 0 Then GoTo Corrupt
        If activeValue Then
            activeCount = activeCount + 1
            If StrComp(roleValue, "Manager", vbTextCompare) = 0 Then managerCount = managerCount + 1
            If StrComp(identityValue, CurrentUser(), vbTextCompare) = 0 Then mSessionRole = roleValue
        End If
    Next lr
    If activeCount = 0 Or managerCount = 0 Then GoTo Corrupt
    If Len(mSessionRole) = 0 Then BlockMutations "This Windows identity is not on the active user allowlist. Managers: " & ManagersList()
    Exit Sub
Corrupt:
    BlockMutations "The bootstrapped user allowlist is empty or corrupt. Access fails closed."
End Sub

Public Sub BlockMutations(ByVal reason As String)
    mMutationBlocked = True
    If Len(mBlockReason) = 0 Then mBlockReason = reason Else mBlockReason = mBlockReason & vbCrLf & reason
End Sub

Public Sub ClearEnvironmentBlock()
    RefreshAuthorization
End Sub

Public Function SessionCanMutate() As Boolean
    If ThisWorkbook.ReadOnly Then
        SessionCanMutate = False
    Else
        SessionCanMutate = (Not mMutationBlocked And Len(mSessionRole) > 0)
    End If
End Function

Private Function IsRecognizedBoolean(ByVal value As Variant) As Boolean
    Dim s As String
    If VarType(value) = vbBoolean Then IsRecognizedBoolean = True: Exit Function
    s = LCase$(Trim$(CStr(value)))
    IsRecognizedBoolean = (s = "true" Or s = "false" Or s = "yes" Or s = "no" Or s = "1" Or s = "0")
End Function

Public Function MutationBlockReason() As String
    If ThisWorkbook.ReadOnly Then
        MutationBlockReason = "Opened read-only - another user is editing. Try later."
    Else
        MutationBlockReason = mBlockReason
    End If
End Function

Public Function RequireMutation() As Boolean
    If Not SessionCanMutate() Then
        MsgBox MutationBlockReason(), vbExclamation, modUtil.APP_NAME & " - Browse mode"
        Exit Function
    End If
    RequireMutation = True
End Function

Public Function RequireRole(ByVal requiredRole As String) As Boolean
    If Not RequireMutation() Then Exit Function
    If StrComp(requiredRole, "Manager", vbTextCompare) = 0 And StrComp(mSessionRole, "Manager", vbTextCompare) <> 0 Then
        MsgBox "This action requires the Manager role.", vbExclamation, modUtil.APP_NAME
        Exit Function
    End If
    RequireRole = True
End Function

Public Function CurrentRole() As String
    CurrentRole = mSessionRole
End Function

Public Function ManagersList() As String
    Dim lo As ListObject
    Dim lr As ListRow
    Dim displayName As String
    On Error GoTo Done
    Set lo = modUtil.GetTable(modUtil.TBL_USERS)
    For Each lr In lo.ListRows
        If modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Active"), False) And StrComp(CStr(modUtil.RowValue(lr, lo, "Role")), "Manager", vbTextCompare) = 0 Then
            displayName = CStr(modUtil.RowValue(lr, lo, "DisplayName"))
            If Len(displayName) = 0 Then displayName = CStr(modUtil.RowValue(lr, lo, "WindowsUsername"))
            If Len(ManagersList) > 0 Then ManagersList = ManagersList & ", "
            ManagersList = ManagersList & displayName
        End If
    Next lr
Done:
End Function
