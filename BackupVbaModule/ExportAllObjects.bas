' ResiPack Database Backup - VBA Export Macro
' This module exports all remaining forms, reports, and queries to text files
' for version control and disaster recovery
'
' USAGE: Add this module to ResiPack.accdb, then run ExportAllRemaining() from the Immediate Window
'        or create a button that calls this subroutine

Option Compare Database
Option Explicit

Public Sub ExportAllObjects()
    '
    ' ExportAllObjects - Complete database backup automation
    ' Exports all 64 database objects: 36 forms, 10 reports, 1 query, 17 modules
    ' This creates a complete text-based backup ready for Git version control
    '
    ' Usage: Press F5 while cursor is in this procedure, or call from Immediate Window
    ' Output: Success/failure message box with summary
    '

    Dim i As Integer
    Dim successCount As Integer
    Dim backupPath As String
    Dim msg As String
    Dim startTime As Double

    On Error GoTo ErrorHandler

    startTime = Timer

    ' Set the backup directory path
    backupPath = CurrentProject.Path & "\Database_Backup"

    ' Verify backup directory exists
    If Dir(backupPath, vbDirectory) = "" Then
        MsgBox "Backup directory not found: " & backupPath, vbCritical
        Exit Sub
    End If

    ' Ensure backup folder structure exists (forms, reports, queries, modules)
    If Not EnsureBackupFolders(backupPath) Then
        MsgBox "Backup cancelled - folder structure not created.", vbExclamation
        Exit Sub
    End If

    successCount = 0

    ' Initialize progress message
    msg = "Starting complete database backup (all 64 objects)..." & vbCrLf & vbCrLf

    ' =========== EXPORT FORMS ==========
    msg = msg & "EXPORTING FORMS (36 total):" & vbCrLf
    successCount = ExportForms(backupPath, msg)

    ' =========== EXPORT REPORTS ==========
    msg = msg & vbCrLf & "EXPORTING REPORTS (10 total):" & vbCrLf
    successCount = successCount + ExportReports(backupPath, msg)

    ' =========== EXPORT QUERIES ==========
    msg = msg & vbCrLf & "EXPORTING QUERIES (1 total):" & vbCrLf
    successCount = successCount + ExportQueries(backupPath, msg)

    ' =========== EXPORT MODULES ==========
    msg = msg & vbCrLf & "EXPORTING MODULES (17 total):" & vbCrLf
    successCount = successCount + ExportModules(backupPath, msg)

    ' Show summary
    msg = msg & vbCrLf & "========================================" & vbCrLf
    msg = msg & "COMPLETE BACKUP FINISHED" & vbCrLf
    msg = msg & "Successfully exported: " & successCount & " of 64 objects" & vbCrLf
    msg = msg & "Time elapsed: " & Format(Timer - startTime, "0.0") & " seconds" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf

    If successCount = 64 Then
        msg = msg & "✓ BACKUP COMPLETE - All objects exported successfully!" & vbCrLf & vbCrLf
    Else
        msg = msg & "⚠ BACKUP INCOMPLETE - " & (64 - successCount) & " objects failed" & vbCrLf & vbCrLf
    End If

    msg = msg & "Next steps:" & vbCrLf
    msg = msg & "1. Close this database" & vbCrLf
    msg = msg & "2. In PowerShell/Git Bash, run:" & vbCrLf
    msg = msg & "   cd C:\GitHub\PalmerElectric\ResiPack" & vbCrLf
    msg = msg & "   git add Database_Backup/" & vbCrLf
    msg = msg & "   git commit -m ""Q3 2026 database backup: All 64 objects""" & vbCrLf
    msg = msg & "   git push" & vbCrLf

    MsgBox msg, IIf(successCount = 64, vbInformation, vbExclamation), "Export Results"

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical

End Sub

Private Function ExportForms(backupPath As String, ByRef msg As String) As Integer

    Dim objForm As AccessObject
    Dim formPath As String
    Dim count As Integer
    Dim formList As Collection

    Set formList = New Collection

    ' Build list of forms to export
    formList.Add "frmInvHeader"
    formList.Add "frmImportVendorPrice"
    formList.Add "frmPartEntry"
    formList.Add "frmOrderEntry"
    formList.Add "frmDeliveryHubEntry"
    formList.Add "frmUserProfile"
    formList.Add "frmGlobalData"
    formList.Add "frmOrderSelect"
    formList.Add "frmInvtIDSelect"
    formList.Add "frmJobOrderMaint"
    formList.Add "frmMfgEntry"
    formList.Add "frmKitEntry"
    formList.Add "frmKitEntryDetail"
    formList.Add "frmOrderItemStatus"
    formList.Add "frmKitExcel"
    formList.Add "frmMainMenu"
    formList.Add "frmKitSelect"
    formList.Add "frmMfgSelect"
    formList.Add "frmPOTranSummary"
    formList.Add "frmPOTranDetailDisplay"
    formList.Add "frmOrderSummary"
    formList.Add "frmOpenOrderSummary"
    formList.Add "frmUserAdmin"
    formList.Add "frmOrderDetail"
    formList.Add "frmPartSelect"
    formList.Add "frmWhseReceiptReturn"
    formList.Add "frmPOTran"
    formList.Add "frmPOTranDetail"
    formList.Add "frmProjectSelect"
    formList.Add "frmResiExport"
    formList.Add "frmSelectPrinter"
    formList.Add "frmSMTP_Config"
    formList.Add "frmBrwsData"
    formList.Add "frmInvDetail"
    formList.Add "frmApexEntry"
    formList.Add "frmPOTranDisplay"

    formPath = backupPath & "\forms"
    count = 0

    Dim i As Integer
    For i = 1 To formList.count
        On Error Resume Next
        formPath = backupPath & "\forms\" & formList(i) & ".txt"
        Application.SaveAsText acForm, formList(i), formPath
        If Err.Number = 0 Then
            msg = msg & "  ✓ " & formList(i) & vbCrLf
            count = count + 1
        Else
            msg = msg & "  ✗ " & formList(i) & " (Error: " & Err.Description & ")" & vbCrLf
        End If
        On Error GoTo 0
    Next i

    msg = msg & "  Total: " & count & "/" & formList.count & " forms exported" & vbCrLf
    ExportForms = count

End Function

Private Function ExportReports(backupPath As String, ByRef msg As String) As Integer

    Dim reportPath As String
    Dim count As Integer
    Dim reportList As Collection

    Set reportList = New Collection

    ' Build list of reports to export
    reportList.Add "rptRecvEmail"
    reportList.Add "rptPurchaseOrderAP"
    reportList.Add "rptEntryApex"
    reportList.Add "rptEntryL"
    reportList.Add "rptOverUnderDetails"
    reportList.Add "rptInvLine"
    reportList.Add "rptGBInvoice"
    reportList.Add "rptPOLine"
    reportList.Add "rptPurchaseOrder"
    reportList.Add "rptUsers"

    reportPath = backupPath & "\reports"
    count = 0

    Dim i As Integer
    For i = 1 To reportList.count
        On Error Resume Next
        reportPath = backupPath & "\reports\" & reportList(i) & ".txt"
        Application.SaveAsText acReport, reportList(i), reportPath
        If Err.Number = 0 Then
            msg = msg & "  ✓ " & reportList(i) & vbCrLf
            count = count + 1
        Else
            msg = msg & "  ✗ " & reportList(i) & " (Error: " & Err.Description & ")" & vbCrLf
        End If
        On Error GoTo 0
    Next i

    msg = msg & "  Total: " & count & "/" & reportList.count & " reports exported" & vbCrLf
    ExportReports = count

End Function

Private Function ExportQueries(backupPath As String, ByRef msg As String) As Integer

    Dim queryPath As String
    Dim count As Integer

    count = 0

    On Error Resume Next
    queryPath = backupPath & "\queries\qry_spPOReportData.txt"
    Application.SaveAsText acQuery, "qry_spPOReportData", queryPath
    If Err.Number = 0 Then
        msg = msg & "  ✓ qry_spPOReportData" & vbCrLf
        count = 1
    Else
        msg = msg & "  ✗ qry_spPOReportData (Error: " & Err.Description & ")" & vbCrLf
    End If
    On Error GoTo 0

    msg = msg & "  Total: " & count & "/1 query exported" & vbCrLf
    ExportQueries = count

End Function

Private Function ExportModules(backupPath As String, ByRef msg As String) As Integer

    Dim moduleList As Collection
    Dim modulePath As String
    Dim count As Integer
    Dim i As Integer
    Dim vbProject As Object
    Dim vbComponent As Object
    Dim codeModule As Object
    Dim moduleCode As String
    Dim outFile As String
    Dim fso As Object

    Set moduleList = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Build list of VBA modules to export
    moduleList.Add "GraybarFunctions"
    moduleList.Add "basBlat"
    moduleList.Add "basBrowseForFolder"
    moduleList.Add "basEnumWindows"
    moduleList.Add "basCallbacks"
    moduleList.Add "basChangeConnection"
    moduleList.Add "basFlashWindow"
    moduleList.Add "basExcelImport"
    moduleList.Add "basGlobalVariables"
    moduleList.Add "basMainMenuLock"
    moduleList.Add "GraybarMain"
    moduleList.Add "basUtility"
    moduleList.Add "cExcelLB"
    moduleList.Add "clsCommonDialog"
    moduleList.Add "MD5Driver"
    moduleList.Add "modReportToPDF"
    moduleList.Add "basExportVBA"

    modulePath = backupPath & "\modules"
    count = 0

    Set vbProject = CurrentProject.VBProject

    For i = 1 To moduleList.count
        On Error Resume Next

        ' Find the module in the VB project
        Set vbComponent = Nothing
        Set vbComponent = vbProject.VBComponents(moduleList(i))

        If Not vbComponent Is Nothing Then
            Set codeModule = vbComponent.CodeModule
            moduleCode = codeModule.Lines(1, codeModule.CountOfLines)

            ' Save to file
            outFile = modulePath & "\" & moduleList(i) & ".bas"
            With fso.CreateTextFile(outFile, True)
                .Write moduleCode
                .Close
            End With

            If Err.Number = 0 Then
                msg = msg & "  ✓ " & moduleList(i) & vbCrLf
                count = count + 1
            Else
                msg = msg & "  ✗ " & moduleList(i) & " (Error: " & Err.Description & ")" & vbCrLf
            End If
        Else
            msg = msg & "  ✗ " & moduleList(i) & " (not found)" & vbCrLf
        End If

        On Error GoTo 0
    Next i

    msg = msg & "  Total: " & count & "/17 modules exported" & vbCrLf
    ExportModules = count

    Set fso = Nothing

End Function

Private Function EnsureBackupFolders(backupPath As String) As Boolean
    '
    ' Verify that all required backup subdirectories exist
    ' If any are missing, prompt user to create them
    ' Returns: True if all folders exist, False if user cancels creation
    '

    Dim fso As Object
    Dim missingFolders As Collection
    Dim folderName As String
    Dim folderPath As String
    Dim confirmMsg As String
    Dim i As Integer

    Set missingFolders = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' List of required subdirectories
    Dim requiredFolders() As String
    requiredFolders = Array("forms", "reports", "queries", "modules")

    ' Check which folders are missing
    For i = LBound(requiredFolders) To UBound(requiredFolders)
        folderPath = backupPath & "\" & requiredFolders(i)
        If Dir(folderPath, vbDirectory) = "" Then
            missingFolders.Add requiredFolders(i)
        End If
    Next i

    ' If no folders missing, we're good
    If missingFolders.count = 0 Then
        EnsureBackupFolders = True
        Exit Function
    End If

    ' Build confirmation message for missing folders
    confirmMsg = "The following backup folders do not exist:" & vbCrLf & vbCrLf

    For i = 1 To missingFolders.count
        confirmMsg = confirmMsg & "  • " & backupPath & "\" & missingFolders(i) & vbCrLf
    Next i

    confirmMsg = confirmMsg & vbCrLf & "Create these folders now?" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "Click YES to create and continue with backup." & vbCrLf
    confirmMsg = confirmMsg & "Click NO to cancel backup."

    ' Prompt user
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Create Backup Folders?") = vbNo Then
        EnsureBackupFolders = False
        Exit Function
    End If

    ' Create missing folders
    On Error Resume Next
    For i = 1 To missingFolders.count
        folderPath = backupPath & "\" & missingFolders(i)
        fso.CreateFolder (folderPath)
        If Err.Number <> 0 Then
            MsgBox "Error creating folder: " & folderPath & vbCrLf & "Error: " & Err.Description, vbCritical
            EnsureBackupFolders = False
            Set fso = Nothing
            Exit Function
        End If
    Next i
    On Error GoTo 0

    Set fso = Nothing
    EnsureBackupFolders = True

End Function
