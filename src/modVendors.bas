Attribute VB_Name = "modVendors"
Option Explicit

Private mBusy As Boolean

Public Function CreateVendor(ByVal vendorName As String, ByVal ownerName As String, ByVal ownerEmail As String, _
    Optional ByVal contactPerson As String = vbNullString, Optional ByVal contactEmail As String = vbNullString, _
    Optional ByVal contactPhone As String = vbNullString, Optional ByVal notes As String = vbNullString, _
    Optional ByVal suppressDuplicateWarning As Boolean = False) As Long
    Dim lo As ListObject
    Dim lr As ListRow
    Dim vendorID As Long
    If mBusy Then Exit Function
    If Not modIdentity.RequireMutation() Then Exit Function
    If Not ValidateVendor(vendorName, ownerName, ownerEmail, contactEmail) Then Exit Function
    If Not suppressDuplicateWarning And VendorNameExists(vendorName) Then
        If MsgBox("A vendor with this name already exists. Add another record anyway?", vbQuestion + vbYesNo, modUtil.APP_NAME) <> vbYes Then Exit Function
    End If
    On Error GoTo Failed
    mBusy = True
    Set lo = modUtil.GetTable(modUtil.TBL_VENDORS)
    vendorID = modUtil.AllocateID("Vendor")
    Set lr = lo.ListRows.Add
    modUtil.SetRowValue lr, lo, "VendorID", vendorID
    modUtil.SetRowValue lr, lo, "VendorName", modUtil.NeutralizeImportedText(Trim$(vendorName))
    modUtil.SetRowValue lr, lo, "BusinessOwnerName", modUtil.NeutralizeImportedText(Trim$(ownerName))
    modUtil.SetRowValue lr, lo, "BusinessOwnerEmail", LCase$(Trim$(ownerEmail))
    modUtil.SetRowValue lr, lo, "VendorContactPerson", modUtil.NeutralizeImportedText(contactPerson)
    modUtil.SetRowValue lr, lo, "VendorContactEmail", LCase$(Trim$(contactEmail))
    modUtil.SetRowValue lr, lo, "VendorContactPhone", modUtil.NeutralizeImportedText(contactPhone)
    modUtil.SetRowValue lr, lo, "Notes", modUtil.NeutralizeImportedText(notes)
    modUtil.SetRowValue lr, lo, "CreatedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, lo, "CreatedOn", Now
    modUtil.SetRowValue lr, lo, "ModifiedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, lo, "ModifiedOn", Now
    modUtil.SetRowValue lr, lo, "Archived", False
    If Not modAudit.AddAudit("CREATE", "Vendor", vendorID, "VendorName=" & vendorName) Then
        lr.Delete
        Err.Raise vbObjectError + 1400, "modVendors.CreateVendor", "The vendor was not created because the audit record failed."
    End If
    CreateVendor = vendorID
    If Not modImport.ImportInProgress() Then modDashboard.RefreshDashboard
Done:
    mBusy = False
    Exit Function
Failed:
    On Error Resume Next
    If Not lr Is Nothing Then lr.Delete
    On Error GoTo 0
    modUtil.ShowError "Vendor", Err.Description
    Resume Done
End Function

Public Function UpdateVendor(ByVal vendorID As Long, ByVal vendorName As String, ByVal ownerName As String, ByVal ownerEmail As String, _
    ByVal contactPerson As String, ByVal contactEmail As String, ByVal contactPhone As String, ByVal notes As String) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    Dim oldValues As Variant
    If mBusy Then Exit Function
    If Not modIdentity.RequireMutation() Then Exit Function
    If Not ValidateVendor(vendorName, ownerName, ownerEmail, contactEmail) Then Exit Function
    Set lo = modUtil.GetTable(modUtil.TBL_VENDORS)
    Set lr = modUtil.FindLongRow(lo, "VendorID", vendorID)
    If lr Is Nothing Then MsgBox "Vendor not found.", vbExclamation, modUtil.APP_NAME: Exit Function
    oldValues = lr.Range.Value2
    On Error GoTo Failed
    mBusy = True
    modUtil.SetRowValue lr, lo, "VendorName", modUtil.NeutralizeImportedText(Trim$(vendorName))
    modUtil.SetRowValue lr, lo, "BusinessOwnerName", modUtil.NeutralizeImportedText(Trim$(ownerName))
    modUtil.SetRowValue lr, lo, "BusinessOwnerEmail", LCase$(Trim$(ownerEmail))
    modUtil.SetRowValue lr, lo, "VendorContactPerson", modUtil.NeutralizeImportedText(contactPerson)
    modUtil.SetRowValue lr, lo, "VendorContactEmail", LCase$(Trim$(contactEmail))
    modUtil.SetRowValue lr, lo, "VendorContactPhone", modUtil.NeutralizeImportedText(contactPhone)
    modUtil.SetRowValue lr, lo, "Notes", modUtil.NeutralizeImportedText(notes)
    modUtil.SetRowValue lr, lo, "ModifiedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, lo, "ModifiedOn", Now
    If Not modAudit.AddAudit("UPDATE", "Vendor", vendorID, "Default owner updated; existing assessment snapshots unchanged.") Then
        lr.Range.Value2 = oldValues
        Err.Raise vbObjectError + 1401, "modVendors.UpdateVendor", "The audit record failed; the vendor was not changed."
    End If
    UpdateVendor = True
    If Not modImport.ImportInProgress() Then modDashboard.RefreshDashboard
Done:
    mBusy = False
    Exit Function
Failed:
    On Error Resume Next
    lr.Range.Value2 = oldValues
    On Error GoTo 0
    modUtil.ShowError "Vendor", Err.Description
    Resume Done
End Function

Private Function ValidateVendor(ByVal vendorName As String, ByVal ownerName As String, ByVal ownerEmail As String, ByVal contactEmail As String) As Boolean
    If Len(Trim$(vendorName)) = 0 Or Len(Trim$(ownerName)) = 0 Or Len(Trim$(ownerEmail)) = 0 Then
        MsgBox "Vendor name, business owner name, and business owner email are required.", vbExclamation, modUtil.APP_NAME
        Exit Function
    End If
    If Not modUtil.IsEmailAddress(ownerEmail) Then MsgBox "Business owner email is invalid.", vbExclamation, modUtil.APP_NAME: Exit Function
    If Len(Trim$(contactEmail)) > 0 And Not modUtil.IsEmailAddress(contactEmail) Then MsgBox "Vendor contact email is invalid.", vbExclamation, modUtil.APP_NAME: Exit Function
    ValidateVendor = True
End Function

Public Function VendorNameExists(ByVal vendorName As String) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    vendorName = modUtil.NeutralizeImportedText(Trim$(vendorName))
    Set lo = modUtil.GetTable(modUtil.TBL_VENDORS)
    For Each lr In lo.ListRows
        If Not modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Archived"), False) Then
            If StrComp(CStr(modUtil.RowValue(lr, lo, "VendorName")), Trim$(vendorName), vbTextCompare) = 0 Then VendorNameExists = True: Exit Function
        End If
    Next lr
End Function

Public Function FindVendorByBusinessKey(ByVal vendorName As String, ByVal ownerEmail As String) As Long
    Dim lo As ListObject
    Dim lr As ListRow
    Set lo = modUtil.GetTable(modUtil.TBL_VENDORS)
    For Each lr In lo.ListRows
        If Not modUtil.ToBoolean(modUtil.RowValue(lr, lo, "Archived"), False) Then
            If StrComp(Trim$(CStr(modUtil.RowValue(lr, lo, "VendorName"))), Trim$(vendorName), vbTextCompare) = 0 _
               And StrComp(Trim$(CStr(modUtil.RowValue(lr, lo, "BusinessOwnerEmail"))), Trim$(ownerEmail), vbTextCompare) = 0 Then
                FindVendorByBusinessKey = CLng(modUtil.RowValue(lr, lo, "VendorID"))
                Exit Function
            End If
        End If
    Next lr
End Function

Public Function ArchiveVendor(ByVal vendorID As Long) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    Dim oldValues As Variant
    If Not modIdentity.RequireRole("Manager") Then Exit Function
    Set lo = modUtil.GetTable(modUtil.TBL_VENDORS)
    Set lr = modUtil.FindLongRow(lo, "VendorID", vendorID)
    If lr Is Nothing Then Exit Function
    oldValues = lr.Range.Value2
    On Error GoTo Failed
    modUtil.SetRowValue lr, lo, "Archived", True
    modUtil.SetRowValue lr, lo, "ArchivedOn", Now
    modUtil.SetRowValue lr, lo, "ArchivedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, lo, "ModifiedBy", modIdentity.CurrentUser()
    modUtil.SetRowValue lr, lo, "ModifiedOn", Now
    If Not modAudit.AddAudit("ARCHIVE", "Vendor", vendorID, "Soft deleted; dependents preserved.") Then lr.Range.Value2 = oldValues: Err.Raise vbObjectError + 1402, , "Audit failed; vendor unchanged."
    ArchiveVendor = True
    Exit Function
Failed:
    On Error Resume Next: lr.Range.Value2 = oldValues: On Error GoTo 0
    modUtil.ShowError "Archive vendor", Err.Description
End Function

Public Function HardDeleteVendor(ByVal vendorID As Long) As Boolean
    Dim lo As ListObject
    Dim assessments As ListObject
    Dim lr As ListRow
    If Not modIdentity.RequireRole("Manager") Then Exit Function
    On Error GoTo BackupFailed
    modBackup.RequireRiskBackup
    Set assessments = modUtil.GetTable(modUtil.TBL_ASSESSMENTS)
    For Each lr In assessments.ListRows
        If CLng(modUtil.RowValue(lr, assessments, "VendorID")) = vendorID Then
            MsgBox "This vendor has dependent assessments and can only be archived.", vbExclamation, modUtil.APP_NAME
            Exit Function
        End If
    Next lr
    Set lo = modUtil.GetTable(modUtil.TBL_VENDORS)
    Set lr = modUtil.FindLongRow(lo, "VendorID", vendorID)
    If lr Is Nothing Then Exit Function
    If Not modAudit.AddAudit("HARD_DELETE", "Vendor", vendorID, "No dependent assessments.") Then MsgBox "Audit failed; nothing deleted.", vbCritical, modUtil.APP_NAME: Exit Function
    lr.Delete
    HardDeleteVendor = True
    Exit Function
BackupFailed:
    modUtil.ShowError "Hard delete vendor", Err.Description
End Function

Public Sub RefreshVendorView()
    Dim src As ListObject
    Dim ws As Worksheet
    Dim lr As ListRow
    Dim outRow As Long
    Set ws = ThisWorkbook.Worksheets("Vendors")
    ws.Range("A5:H1048576").ClearContents
    ws.Range("A4:H4").Value = Array("Vendor ID", "Vendor Name", "Business Owner", "Owner Email", "Contact", "Contact Email", "Phone", "Notes")
    outRow = 5
    Set src = modUtil.GetTable(modUtil.TBL_VENDORS)
    For Each lr In src.ListRows
        If Not modUtil.ToBoolean(modUtil.RowValue(lr, src, "Archived"), False) Then
            ws.Cells(outRow, 1).Value = modUtil.RowValue(lr, src, "VendorID")
            ws.Cells(outRow, 2).Value = modUtil.RowValue(lr, src, "VendorName")
            ws.Cells(outRow, 3).Value = modUtil.RowValue(lr, src, "BusinessOwnerName")
            ws.Cells(outRow, 4).Value = modUtil.RowValue(lr, src, "BusinessOwnerEmail")
            ws.Cells(outRow, 5).Value = modUtil.RowValue(lr, src, "VendorContactPerson")
            ws.Cells(outRow, 6).Value = modUtil.RowValue(lr, src, "VendorContactEmail")
            ws.Cells(outRow, 7).NumberFormat = "@"
            ws.Cells(outRow, 7).Value = modUtil.RowValue(lr, src, "VendorContactPhone")
            ws.Cells(outRow, 8).Value = modUtil.RowValue(lr, src, "Notes")
            outRow = outRow + 1
        End If
    Next lr
End Sub
