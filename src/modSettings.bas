Attribute VB_Name = "modSettings"
Option Explicit

Public Function GetSetting(ByVal settingName As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim lo As ListObject
    Dim lr As ListRow
    On Error GoTo Missing
    Set lo = modUtil.GetTable(modUtil.TBL_SETTINGS)
    Set lr = modUtil.FindTextRow(lo, "SettingName", settingName)
    If lr Is Nothing Then GoTo Missing
    GetSetting = CStr(modUtil.RowValue(lr, lo, "SettingValue"))
    Exit Function
Missing:
    GetSetting = defaultValue
End Function

Public Function GetLong(ByVal settingName As String, ByVal defaultValue As Long) As Long
    Dim raw As String
    raw = GetSetting(settingName, CStr(defaultValue))
    If IsNumeric(raw) Then GetLong = CLng(raw) Else GetLong = defaultValue
End Function

Public Sub SetSystemSetting(ByVal settingName As String, ByVal settingValue As String)
    Dim lo As ListObject
    Dim lr As ListRow
    Set lo = modUtil.GetTable(modUtil.TBL_SETTINGS)
    Set lr = modUtil.FindTextRow(lo, "SettingName", settingName)
    If lr Is Nothing Then Err.Raise vbObjectError + 1100, "modSettings.SetSystemSetting", "Unknown setting: " & settingName
    modUtil.SetRowValue lr, lo, "SettingValue", settingValue
End Sub

Public Function UpdateSetting(ByVal settingName As String, ByVal settingValue As String) As Boolean
    Dim oldValue As String
    Dim storedValue As String
    If Not modIdentity.RequireRole("Manager") Then Exit Function
    If Not ValidateSetting(settingName, settingValue) Then Exit Function
    If settingName = "EmailMode" And StrComp(settingValue, "Direct", vbTextCompare) = 0 Then
        If MsgBox("Direct mode calls Outlook .Send after the batch preview. Events remain Queued until GUID reconciliation finds Sent Items. Enable Direct mode?", vbCritical + vbYesNo + vbDefaultButton2, modUtil.APP_NAME) <> vbYes Then Exit Function
    End If
    storedValue = settingValue
    Select Case settingName
        Case "OrgName", "EmailInitialSubject", "EmailInitialTemplate", "EmailFollowupSubject", "EmailFollowupTemplate", "SendingAccountSMTP", "InternalDomainAllowlist"
            storedValue = modUtil.NeutralizeImportedText(settingValue)
    End Select
    oldValue = GetSetting(settingName, vbNullString)
    On Error GoTo Failed
    SetSystemSetting settingName, storedValue
    If Not modAudit.AddAudit("UPDATE_SETTING", "Setting", settingName, "Old=" & oldValue & "; New=" & settingValue) Then
        SetSystemSetting settingName, oldValue
        Err.Raise vbObjectError + 1101, "modSettings.UpdateSetting", "The audit record could not be written; the setting was not changed."
    End If
    UpdateSetting = True
    Exit Function
Failed:
    modUtil.ShowError "Settings", Err.Description
End Function

Private Function ValidateSetting(ByVal settingName As String, ByVal settingValue As String) As Boolean
    Select Case settingName
        Case "FollowupDays", "BatchCap", "StaleBackupKeep", "ReconcileGraceHours"
            If Not IsNumeric(settingValue) Or CLng(settingValue) < 1 Then
                MsgBox settingName & " must be a positive whole number.", vbExclamation, modUtil.APP_NAME
                Exit Function
            End If
        Case "EmailMode"
            If StrComp(settingValue, "Draft", vbTextCompare) <> 0 And StrComp(settingValue, "Direct", vbTextCompare) <> 0 Then
                MsgBox "EmailMode must be Draft or Direct.", vbExclamation, modUtil.APP_NAME
                Exit Function
            End If
        Case "InternalDomainAllowlist"
            If Len(Trim$(settingValue)) = 0 Then
                MsgBox "At least one exact internal email domain is required.", vbExclamation, modUtil.APP_NAME
                Exit Function
            End If
        Case "BootstrapState", "NextVendorID", "NextAssessmentID", "NextEventID", "SchemaVersion", "LastBackupDate"
            MsgBox "This setting is managed by the application.", vbExclamation, modUtil.APP_NAME
            Exit Function
    End Select
    ValidateSetting = True
End Function
