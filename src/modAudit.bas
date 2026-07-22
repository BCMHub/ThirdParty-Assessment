Attribute VB_Name = "modAudit"
Option Explicit

Public Function AddAudit(ByVal actionName As String, ByVal entityType As String, ByVal entityID As Variant, ByVal detail As String) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    On Error GoTo Failed
    Set lo = modUtil.GetTable(modUtil.TBL_AUDIT)
    Set lr = lo.ListRows.Add
    modUtil.SetRowValue lr, lo, "Timestamp", Now
    modUtil.SetRowValue lr, lo, "WindowsUsername", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, lo, "Action", actionName
    modUtil.SetRowValue lr, lo, "EntityType", entityType
    modUtil.SetRowValue lr, lo, "EntityID", CStr(entityID)
    modUtil.SetRowValue lr, lo, "Detail", detail
    AddAudit = True
    Exit Function
Failed:
    On Error Resume Next
    If Not lr Is Nothing Then lr.Delete
    On Error GoTo 0
    AddAudit = False
End Function

Public Sub AuditOrRaise(ByVal actionName As String, ByVal entityType As String, ByVal entityID As Variant, ByVal detail As String)
    If Not AddAudit(actionName, entityType, entityID, detail) Then
        Err.Raise vbObjectError + 1200, "modAudit.AuditOrRaise", "The required audit record could not be written."
    End If
End Sub
