Attribute VB_Name = "modEmail"
Option Explicit

Private Const OL_FOLDER_SENT_MAIL As Long = 5
Private Const OL_FOLDER_OUTBOX As Long = 4
Private Const OL_FOLDER_DRAFTS As Long = 16
Private Const OL_MAIL_ITEM As Long = 0
Private Const OL_TEXT As Long = 1
Private Const CORRELATION_PROPERTY As String = "TPATCorrelationToken"
Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
Private mBusy As Boolean
Private mSummaryValid As Boolean
Private mFirstSent As Object
Private mLastSentKind As Object
Private mLastSentAny As Object
Private mFollowupCounts As Object
Private mActiveEvents As Object

Public Function PrepareEmail(ByVal assessmentID As Long, ByVal kind As String, Optional ByVal forceDirect As Boolean = False, Optional ByVal directAlreadyConfirmed As Boolean = False) As Long
    Dim assessments As ListObject
    Dim assessmentRow As ListRow
    Dim events As ListObject
    Dim eventRow As ListRow
    Dim eventID As Long
    Dim token As String
    Dim ownerName As String
    Dim recipient As String
    Dim vendorName As String
    Dim cycle As String
    Dim subjectText As String
    Dim bodyText As String
    Dim outlookApp As Object
    Dim mailItem As Object
    Dim tokenProperty As Object
    Dim account As Object
    Dim directMode As Boolean
    Dim storeID As String
    Dim sendingAccount As String
    Dim errText As String
    If mBusy Then Exit Function
    If Not modIdentity.RequireMutation() Then Exit Function
    kind = CanonicalKind(kind)
    If Len(kind) = 0 Then MsgBox "Email kind must be Initial or Followup.", vbExclamation, modUtil.APP_NAME: Exit Function
    Set assessments = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    Set assessmentRow = modUtil.FindLongRow(assessments, "AssessmentID", assessmentID)
    If assessmentRow Is Nothing Then MsgBox "Assessment not found.", vbExclamation, modUtil.APP_NAME: Exit Function
    If modUtil.ToBoolean(modUtil.RowValue(assessmentRow, assessments, "Archived"), False) Then MsgBox "Archived assessments cannot create email.", vbExclamation, modUtil.APP_NAME: Exit Function
    If kind = "Followup" And Not HasConfirmedInitial(assessmentID) Then MsgBox "Follow-up is blocked until an Initial event is confirmed Sent.", vbExclamation, modUtil.APP_NAME: Exit Function
    If HasActiveEvent(assessmentID, kind) Then MsgBox "A pending or unresolved " & kind & " email already exists for this assessment. Reconcile or cancel it before retrying.", vbExclamation, modUtil.APP_NAME: Exit Function
    If Not modAssessments.GetAssessmentOwner(assessmentID, ownerName, recipient, vendorName, cycle) Then MsgBox "Assessment owner snapshot could not be resolved.", vbExclamation, modUtil.APP_NAME: Exit Function
    If Not modUtil.IsEmailAddress(recipient) Then MsgBox "Assessment owner email is invalid.", vbExclamation, modUtil.APP_NAME: Exit Function
    If Not modUtil.IsAllowedInternalEmail(recipient) Then MsgBox "Recipient domain is not an exact match in InternalDomainAllowlist.", vbCritical, modUtil.APP_NAME: Exit Function
    directMode = forceDirect Or StrComp(modSettings.GetSetting("EmailMode", "Draft"), "Direct", vbTextCompare) = 0
    If directMode And StrComp(modIdentity.CurrentRole(), "Manager", vbTextCompare) <> 0 Then MsgBox "Direct-send mode requires Manager role.", vbExclamation, modUtil.APP_NAME: Exit Function
    If directMode And Not directAlreadyConfirmed Then
        If MsgBox("DIRECT SEND recipient:" & vbCrLf & recipient & vbCrLf & vbCrLf & "This one item is within BatchCap=" & modSettings.GetLong("BatchCap", 50) & ". Outlook .Send will record Queued until reconciliation. Continue?", vbCritical + vbYesNo + vbDefaultButton2, modUtil.APP_NAME) <> vbYes Then Exit Function
    End If
    On Error GoTo FailedBeforeEvent
    subjectText = modUtil.RenderTemplate(modSettings.GetSetting(IIf(kind = "Initial", "EmailInitialSubject", "EmailFollowupSubject")), vendorName, ownerName, cycle)
    bodyText = modUtil.RenderTemplate(modSettings.GetSetting(IIf(kind = "Initial", "EmailInitialTemplate", "EmailFollowupTemplate")), vendorName, ownerName, cycle)
    token = modUtil.NewGuidToken()
    eventID = modUtil.AllocateID("Event")
    Set events = modUtil.GetTable(modUtil.TBL_EMAIL_EVENTS)
    Set eventRow = events.ListRows.Add
    modUtil.SetRowValue eventRow, events, "EventID", eventID
    modUtil.SetRowValue eventRow, events, "AssessmentID", assessmentID
    modUtil.SetRowValue eventRow, events, "Kind", kind
    modUtil.SetRowValue eventRow, events, "State", "Prepared"
    modUtil.SetRowValue eventRow, events, "CorrelationToken", token
    modUtil.SetRowValue eventRow, events, "Recipient", recipient
    modUtil.SetRowValue eventRow, events, "Subject", modUtil.NeutralizeImportedText(subjectText)
    modUtil.SetRowValue eventRow, events, "PreparedOn", Now
    modUtil.SetRowValue eventRow, events, "PreparedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue eventRow, events, "LastStateOn", Now
    If Not modAudit.AddAudit("EMAIL_PREPARED", "EmailEvent", eventID, "AssessmentID=" & assessmentID & "; Kind=" & kind & "; GUID=" & token) Then
        eventRow.Delete
        Err.Raise vbObjectError + 1700, , "The email intent was not persisted because its audit record failed. Outlook was not called."
    End If
    InvalidateSummaryCache
    mBusy = True
    On Error GoTo OutlookFailed
    Set outlookApp = CreateObject("Outlook.Application")
    Set account = ResolveSendingAccount(outlookApp)
    If account Is Nothing Then Err.Raise vbObjectError + 1701, , "No matching Outlook sending account is configured."
    Set mailItem = outlookApp.CreateItem(OL_MAIL_ITEM)
    Set mailItem.SendUsingAccount = account
    mailItem.To = recipient
    mailItem.Subject = subjectText
    mailItem.Body = bodyText
    Set tokenProperty = mailItem.UserProperties.Add(CORRELATION_PROPERTY, OL_TEXT, True)
    tokenProperty.Value = token
    mailItem.Save
    storeID = CStr(mailItem.Parent.StoreID)
    modUtil.SetRowValue eventRow, events, "SendingAccount", AccountAddress(account)
    modUtil.SetRowValue eventRow, events, "DraftEntryID", CStr(mailItem.EntryID)
    modUtil.SetRowValue eventRow, events, "StoreID", storeID
    If Not ChangeEventState(eventID, "DraftCreated", "Draft saved with GUID user property.") Then Err.Raise vbObjectError + 1702, , "Draft was created but state audit failed. Reconcile by GUID before retrying."
    If directMode Then
        mailItem.Send
        If Not ChangeEventState(eventID, "Queued", "Outlook .Send invoked; queued is not delivery proof.") Then Err.Raise vbObjectError + 1704, , "Outlook queued the item but the state audit failed. Reconcile before retrying."
    End If
    PrepareEmail = eventID
    modDashboard.RefreshDashboard
Done:
    mBusy = False
    Exit Function
OutlookFailed:
    errText = Err.Description
    On Error Resume Next
    If mailItem Is Nothing Or Len(CStr(modUtil.RowValue(eventRow, events, "DraftEntryID"))) = 0 Then
        ChangeEventState eventID, "Failed", "Outlook error before a draft EntryID was captured: " & errText
    Else
        modAudit.AddAudit "EMAIL_ERROR", "EmailEvent", eventID, "Tracked draft retained in active state for GUID reconciliation: " & errText
    End If
    On Error GoTo 0
    modUtil.ShowError "Email", errText & vbCrLf & "The persisted event remains tracked. Reconcile by GUID before retrying."
    Resume Done
FailedBeforeEvent:
    If Not eventRow Is Nothing Then
        On Error Resume Next: eventRow.Delete: On Error GoTo 0
    End If
    modUtil.ShowError "Email", Err.Description
    Resume Done
End Function

Public Function HasActiveEvent(ByVal assessmentID As Long, ByVal kind As String) As Boolean
    EnsureSummaryCache
    HasActiveEvent = mActiveEvents.Exists(CStr(assessmentID) & "|" & CanonicalKind(kind))
End Function

Public Function HasConfirmedInitial(ByVal assessmentID As Long) As Boolean
    HasConfirmedInitial = IsDate(FirstSentDate(assessmentID, "Initial"))
End Function

Public Function FirstSentDate(ByVal assessmentID As Long, ByVal kind As String) As Variant
    Dim key As String
    EnsureSummaryCache
    key = CStr(assessmentID) & "|" & CanonicalKind(kind)
    If mFirstSent.Exists(key) Then FirstSentDate = CDate(mFirstSent(key)) Else FirstSentDate = Empty
End Function

Public Function LastSentDate(ByVal assessmentID As Long, Optional ByVal kind As String = vbNullString) As Variant
    Dim key As String
    EnsureSummaryCache
    If Len(kind) = 0 Then
        key = CStr(assessmentID)
        If mLastSentAny.Exists(key) Then LastSentDate = CDate(mLastSentAny(key)) Else LastSentDate = Empty
    Else
        key = CStr(assessmentID) & "|" & CanonicalKind(kind)
        If mLastSentKind.Exists(key) Then LastSentDate = CDate(mLastSentKind(key)) Else LastSentDate = Empty
    End If
End Function

Public Function FollowupCount(ByVal assessmentID As Long) As Long
    EnsureSummaryCache
    If mFollowupCounts.Exists(CStr(assessmentID)) Then FollowupCount = CLng(mFollowupCounts(CStr(assessmentID)))
End Function

Public Function ChangeEventState(ByVal eventID As Long, ByVal newState As String, ByVal detail As String, Optional ByVal sentOnValue As Variant, Optional ByVal internetMessageID As String = vbNullString) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    Dim oldState As String
    Dim oldRow As Variant
    If Not IsValidEventState(newState) Then Exit Function
    Set lo = modUtil.GetTable(modUtil.TBL_EMAIL_EVENTS)
    Set lr = modUtil.FindLongRow(lo, "EventID", eventID)
    If lr Is Nothing Then Exit Function
    oldState = CStr(modUtil.RowValue(lr, lo, "State"))
    If Not IsAllowedEventTransition(oldState, newState) Then Exit Function
    If oldState = newState And Len(internetMessageID) = 0 Then ChangeEventState = True: Exit Function
    oldRow = lr.Range.Value2
    On Error GoTo Failed
    modUtil.SetRowValue lr, lo, "State", newState
    modUtil.SetRowValue lr, lo, "LastStateOn", Now
    modUtil.SetRowValue lr, lo, "ErrorDetail", IIf(newState = "Failed" Or newState = "Unresolved", detail, vbNullString)
    If newState = "Sent" Then
        If IsMissing(sentOnValue) Or Not IsDate(sentOnValue) Then modUtil.SetRowValue lr, lo, "SentOn", Now Else modUtil.SetRowValue lr, lo, "SentOn", CDate(sentOnValue)
        If Len(internetMessageID) > 0 Then modUtil.SetRowValue lr, lo, "InternetMessageID", internetMessageID
    End If
    If Not modAudit.AddAudit("EMAIL_STATE", "EmailEvent", eventID, oldState & " -> " & newState & "; " & detail) Then
        lr.Range.Value2 = oldRow
        Exit Function
    End If
    InvalidateSummaryCache
    ChangeEventState = True
    Exit Function
Failed:
    On Error Resume Next: lr.Range.Value2 = oldRow: On Error GoTo 0
End Function

Public Sub Reconcile(Optional ByVal showSummary As Boolean = True)
    Dim outlookApp As Object
    Dim ns As Object
    Dim lo As ListObject
    Dim lr As ListRow
    Dim eventID As Long
    Dim stateValue As String
    Dim token As String
    Dim storeID As String
    Dim sendingAccount As String
    Dim foundItem As Object
    Dim folderKind As String
    Dim matchCount As Long
    Dim checked As Long
    Dim resolved As Long
    Dim graceHours As Long
    Dim preparedOn As Variant
    Dim messageID As String
    If Not modIdentity.RequireMutation() Then Exit Sub
    On Error GoTo Failed
    Set outlookApp = CreateObject("Outlook.Application")
    Set ns = outlookApp.GetNamespace("MAPI")
    graceHours = modSettings.GetLong("ReconcileGraceHours", 24)
    Set lo = modUtil.GetTable(modUtil.TBL_EMAIL_EVENTS)
    For Each lr In lo.ListRows
        stateValue = CStr(modUtil.RowValue(lr, lo, "State"))
        If stateValue = "Prepared" Or stateValue = "DraftCreated" Or stateValue = "Queued" Or stateValue = "Unresolved" Then
            checked = checked + 1
            eventID = CLng(modUtil.RowValue(lr, lo, "EventID"))
            token = CStr(modUtil.RowValue(lr, lo, "CorrelationToken"))
            storeID = CStr(modUtil.RowValue(lr, lo, "StoreID"))
            sendingAccount = CStr(modUtil.RowValue(lr, lo, "SendingAccount"))
            Set foundItem = Nothing: folderKind = vbNullString: matchCount = 0
            FindByGuid ns, token, storeID, sendingAccount, foundItem, folderKind, matchCount
            If matchCount = 1 Then
                Select Case folderKind
                    Case "Drafts"
                        If stateValue = "Prepared" Or stateValue = "Unresolved" Then If ChangeEventState(eventID, "DraftCreated", "Reconciled by GUID in Drafts.") Then resolved = resolved + 1
                    Case "Outbox"
                        If ChangeEventState(eventID, "Queued", "Reconciled by GUID in Outbox.") Then resolved = resolved + 1
                    Case "Sent"
                        messageID = ReadInternetMessageID(foundItem)
                        If ChangeEventState(eventID, "Sent", "Exactly one GUID match in Sent Items.", foundItem.SentOn, messageID) Then resolved = resolved + 1
                End Select
            ElseIf matchCount > 1 Then
                If ChangeEventState(eventID, "Unresolved", "Multiple GUID matches found; manual review required.") Then resolved = resolved + 1
            Else
                preparedOn = modUtil.RowValue(lr, lo, "PreparedOn")
                If IsDate(preparedOn) Then
                    If DateDiff("h", CDate(preparedOn), Now) >= graceHours Then
                        If ChangeEventState(eventID, "Unresolved", "No GUID match after " & graceHours & " hour grace period; item may be moved/archived or Sent-copy saving disabled.") Then resolved = resolved + 1
                    End If
                End If
            End If
        End If
    Next lr
    modDashboard.RefreshDashboard
    If showSummary Then MsgBox "Reconciliation checked " & checked & " event(s) and changed " & resolved & ".", vbInformation, modUtil.APP_NAME
    Exit Sub
Failed:
    modUtil.ShowError "Reconcile", "Classic Outlook/MAPI reconciliation failed: " & Err.Description
End Sub

Private Sub FindByGuid(ByVal ns As Object, ByVal token As String, ByVal preferredStoreID As String, ByVal sendingAccount As String, ByRef foundItem As Object, ByRef folderKind As String, ByRef matchCount As Long)
    Dim stores As Object
    Dim store As Object
    Dim account As Object
    Dim accountStoreID As String
    Dim searched As Object
    Set stores = ns.Stores
    Set searched = CreateObject("Scripting.Dictionary")
    If Len(preferredStoreID) = 0 And Len(sendingAccount) = 0 Then
        For Each store In stores
            SearchStore store, token, foundItem, folderKind, matchCount
        Next store
        Exit Sub
    End If
    For Each store In stores
        If StrComp(CStr(store.StoreID), preferredStoreID, vbBinaryCompare) = 0 Then
            SearchStore store, token, foundItem, folderKind, matchCount
            searched(CStr(store.StoreID)) = True
        End If
    Next store
    If Len(sendingAccount) > 0 Then
        For Each account In ns.Accounts
            If StrComp(AccountAddress(account), sendingAccount, vbTextCompare) = 0 Then
                On Error Resume Next
                accountStoreID = CStr(account.DeliveryStore.StoreID)
                On Error GoTo 0
                If Len(accountStoreID) > 0 And Not searched.Exists(accountStoreID) Then
                    For Each store In stores
                        If StrComp(CStr(store.StoreID), accountStoreID, vbBinaryCompare) = 0 Then SearchStore store, token, foundItem, folderKind, matchCount
                    Next store
                End If
            End If
        Next account
    End If
End Sub

Private Sub SearchStore(ByVal store As Object, ByVal token As String, ByRef foundItem As Object, ByRef folderKind As String, ByRef matchCount As Long)
    Dim folder As Object
    On Error Resume Next
    Set folder = store.GetDefaultFolder(OL_FOLDER_DRAFTS)
    On Error GoTo 0
    If Not folder Is Nothing Then SearchFolder folder, token, "Drafts", foundItem, folderKind, matchCount
    Set folder = Nothing
    On Error Resume Next
    Set folder = store.GetDefaultFolder(OL_FOLDER_OUTBOX)
    On Error GoTo 0
    If Not folder Is Nothing Then SearchFolder folder, token, "Outbox", foundItem, folderKind, matchCount
    Set folder = Nothing
    On Error Resume Next
    Set folder = store.GetDefaultFolder(OL_FOLDER_SENT_MAIL)
    On Error GoTo 0
    If Not folder Is Nothing Then SearchFolder folder, token, "Sent", foundItem, folderKind, matchCount
End Sub

Private Sub SearchFolder(ByVal folder As Object, ByVal token As String, ByVal kind As String, ByRef foundItem As Object, ByRef folderKind As String, ByRef matchCount As Long)
    Dim items As Object
    Dim item As Object
    Dim customProperty As Object
    On Error Resume Next
    Set items = folder.Items
    For Each item In items
        Set customProperty = Nothing
        Set customProperty = item.UserProperties(CORRELATION_PROPERTY)
        If Not customProperty Is Nothing Then
            If StrComp(CStr(customProperty.Value), token, vbTextCompare) = 0 Then
                matchCount = matchCount + 1
                Set foundItem = item
                folderKind = kind
            End If
        End If
    Next item
    On Error GoTo 0
End Sub

Private Function ResolveSendingAccount(ByVal outlookApp As Object) As Object
    Dim account As Object
    Dim desired As String
    desired = Trim$(modSettings.GetSetting("SendingAccountSMTP", vbNullString))
    For Each account In outlookApp.Session.Accounts
        If Len(desired) = 0 Or StrComp(AccountAddress(account), desired, vbTextCompare) = 0 Then Set ResolveSendingAccount = account: Exit Function
    Next account
End Function

Private Function AccountAddress(ByVal account As Object) As String
    On Error Resume Next
    AccountAddress = CStr(account.SmtpAddress)
    If Len(AccountAddress) = 0 Then AccountAddress = CStr(account.DisplayName)
    On Error GoTo 0
End Function

Private Function ReadInternetMessageID(ByVal mailItem As Object) As String
    On Error Resume Next
    ReadInternetMessageID = CStr(mailItem.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID))
    On Error GoTo 0
End Function

Private Function CanonicalKind(ByVal kind As String) As String
    If StrComp(kind, "Initial", vbTextCompare) = 0 Then CanonicalKind = "Initial"
    If StrComp(kind, "Followup", vbTextCompare) = 0 Or StrComp(kind, "Follow-up", vbTextCompare) = 0 Then CanonicalKind = "Followup"
End Function

Private Function IsValidEventState(ByVal stateValue As String) As Boolean
    Select Case stateValue
        Case "Prepared", "DraftCreated", "Queued", "Sent", "Unresolved", "Failed", "Cancelled": IsValidEventState = True
    End Select
End Function

Private Function IsAllowedEventTransition(ByVal oldState As String, ByVal newState As String) As Boolean
    If oldState = newState Then IsAllowedEventTransition = True: Exit Function
    Select Case oldState
        Case "Prepared": IsAllowedEventTransition = (newState = "DraftCreated" Or newState = "Queued" Or newState = "Sent" Or newState = "Unresolved" Or newState = "Failed" Or newState = "Cancelled")
        Case "DraftCreated": IsAllowedEventTransition = (newState = "Queued" Or newState = "Sent" Or newState = "Unresolved" Or newState = "Failed" Or newState = "Cancelled")
        Case "Queued": IsAllowedEventTransition = (newState = "Sent" Or newState = "Unresolved" Or newState = "Failed" Or newState = "Cancelled")
        Case "Unresolved": IsAllowedEventTransition = (newState = "DraftCreated" Or newState = "Queued" Or newState = "Sent" Or newState = "Failed" Or newState = "Cancelled")
    End Select
End Function

Public Function CancelEvent(ByVal eventID As Long, ByVal reason As String) As Boolean
    If Not modIdentity.RequireMutation() Then Exit Function
    CancelEvent = ChangeEventState(eventID, "Cancelled", reason)
End Function

Public Sub InvalidateSummaryCache()
    mSummaryValid = False
End Sub

Private Sub EnsureSummaryCache()
    Dim lo As ListObject
    Dim lr As ListRow
    Dim assessmentKey As String
    Dim kindKey As String
    Dim stateValue As String
    Dim sentValue As Variant
    Dim serialValue As Double
    If mSummaryValid Then Exit Sub
    Set mFirstSent = CreateObject("Scripting.Dictionary")
    Set mLastSentKind = CreateObject("Scripting.Dictionary")
    Set mLastSentAny = CreateObject("Scripting.Dictionary")
    Set mFollowupCounts = CreateObject("Scripting.Dictionary")
    Set mActiveEvents = CreateObject("Scripting.Dictionary")
    Set lo = modUtil.GetTable(modUtil.TBL_EMAIL_EVENTS)
    For Each lr In lo.ListRows
        assessmentKey = CStr(modUtil.RowValue(lr, lo, "AssessmentID"))
        kindKey = assessmentKey & "|" & CanonicalKind(CStr(modUtil.RowValue(lr, lo, "Kind")))
        stateValue = CStr(modUtil.RowValue(lr, lo, "State"))
        If stateValue = "Prepared" Or stateValue = "DraftCreated" Or stateValue = "Queued" Or stateValue = "Unresolved" Then mActiveEvents(kindKey) = True
        If stateValue = "Sent" Then
            sentValue = modUtil.RowValue(lr, lo, "SentOn")
            If IsDate(sentValue) Then
                serialValue = CDbl(CDate(sentValue))
                If Not mFirstSent.Exists(kindKey) Then
                    mFirstSent(kindKey) = serialValue
                ElseIf serialValue < CDbl(mFirstSent(kindKey)) Then
                    mFirstSent(kindKey) = serialValue
                End If
                If Not mLastSentKind.Exists(kindKey) Then
                    mLastSentKind(kindKey) = serialValue
                ElseIf serialValue > CDbl(mLastSentKind(kindKey)) Then
                    mLastSentKind(kindKey) = serialValue
                End If
                If Not mLastSentAny.Exists(assessmentKey) Then
                    mLastSentAny(assessmentKey) = serialValue
                ElseIf serialValue > CDbl(mLastSentAny(assessmentKey)) Then
                    mLastSentAny(assessmentKey) = serialValue
                End If
                If Right$(kindKey, 9) = "|Followup" Then
                    If Not mFollowupCounts.Exists(assessmentKey) Then mFollowupCounts(assessmentKey) = 1 Else mFollowupCounts(assessmentKey) = CLng(mFollowupCounts(assessmentKey)) + 1
                End If
            End If
        End If
    Next lr
    mSummaryValid = True
End Sub
