Attribute VB_Name = "modBackup"
Option Explicit

Public Function EnsureDailyBackup(Optional ByVal forceBackup As Boolean = False, Optional ByVal showFailure As Boolean = True) As Boolean
    Dim backupFolder As String
    Dim datePrefix As String
    Dim backupPath As String
    Dim baseName As String
    On Error GoTo Failed
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 1300, "modBackup.EnsureDailyBackup", "Save the workbook before creating a backup."
    backupFolder = ThisWorkbook.Path & Application.PathSeparator & "Backups"
    If Dir$(backupFolder, vbDirectory) = vbNullString Then MkDir backupFolder
    baseName = modUtil.SafeFileBaseName(ThisWorkbook.Name)
    datePrefix = baseName & "_" & Format$(Date, "yyyy-mm-dd") & "_"
    If Not forceBackup Then
        If Len(Dir$(backupFolder & Application.PathSeparator & datePrefix & "*.xlsm")) > 0 Then
            EnsureDailyBackup = True
            Exit Function
        End If
    End If
    backupPath = backupFolder & Application.PathSeparator & datePrefix & Format$(Time, "hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs backupPath
    If Not modAudit.AddAudit(IIf(forceBackup, "BACKUP_FORCED", "BACKUP_DAILY"), "Workbook", ThisWorkbook.Name, backupPath) Then
        Err.Raise vbObjectError + 1301, "modBackup.EnsureDailyBackup", "Backup was created, but its audit record could not be written."
    End If
    PruneBackups backupFolder, baseName, modSettings.GetLong("StaleBackupKeep", 10)
    EnsureDailyBackup = True
    Exit Function
Failed:
    EnsureDailyBackup = False
    If showFailure Then MsgBox "Backup failed: " & Err.Description, vbCritical, modUtil.APP_NAME
End Function

Public Sub RequireRiskBackup()
    If Not modIdentity.RequireMutation() Then Err.Raise vbObjectError + 1302, "modBackup.RequireRiskBackup", modIdentity.MutationBlockReason()
    If Not EnsureDailyBackup(True, True) Then
        Err.Raise vbObjectError + 1303, "modBackup.RequireRiskBackup", "The high-risk operation was blocked because its required backup failed."
    End If
End Sub

Private Sub PruneBackups(ByVal backupFolder As String, ByVal baseName As String, ByVal keepCount As Long)
    Dim names() As String
    Dim dates() As Date
    Dim count As Long
    Dim fileName As String
    Dim i As Long
    Dim j As Long
    Dim swapName As String
    Dim swapDate As Date
    If keepCount < 1 Then keepCount = 1
    fileName = Dir$(backupFolder & Application.PathSeparator & baseName & "_*.xlsm")
    Do While Len(fileName) > 0
        count = count + 1
        ReDim Preserve names(1 To count)
        ReDim Preserve dates(1 To count)
        names(count) = fileName
        dates(count) = FileDateTime(backupFolder & Application.PathSeparator & fileName)
        fileName = Dir$()
    Loop
    For i = 1 To count - 1
        For j = i + 1 To count
            If dates(j) > dates(i) Then
                swapDate = dates(i): dates(i) = dates(j): dates(j) = swapDate
                swapName = names(i): names(i) = names(j): names(j) = swapName
            End If
        Next j
    Next i
    On Error Resume Next
    For i = keepCount + 1 To count
        Kill backupFolder & Application.PathSeparator & names(i)
    Next i
    On Error GoTo 0
End Sub

Public Function TestBackupFolder(ByRef failureReason As String) As Boolean
    Dim folderPath As String
    Dim testPath As String
    Dim fileNumber As Integer
    Dim errorDescription As String
    On Error GoTo Failed
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise 5, , "Workbook has no saved path."
    folderPath = ThisWorkbook.Path & Application.PathSeparator & "Backups"
    If Dir$(folderPath, vbDirectory) = vbNullString Then MkDir folderPath
    testPath = folderPath & Application.PathSeparator & ".tpat_write_test_" & Format$(Now, "yyyymmddhhnnss") & ".tmp"
    fileNumber = FreeFile
    Open testPath For Output As #fileNumber
    Print #fileNumber, "write-test"
    Close #fileNumber
    Kill testPath
    TestBackupFolder = True
    Exit Function
Failed:
    errorDescription = Err.Description
    On Error Resume Next
    Close #fileNumber
    If Len(testPath) > 0 And Len(Dir$(testPath)) > 0 Then Kill testPath
    failureReason = errorDescription
End Function
