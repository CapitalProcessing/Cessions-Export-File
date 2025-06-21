Attribute VB_Name = "Module2"
' === MOMENT-ENHANCED VBA MODULE ===
' Key "moments" added to ProcessItems and ProcessSingle for Excel stability after SaveAlteredWorkbook and before summary operations

Option Explicit
Sub ProcessAllTemplates()
    Dim wsAuto As Worksheet
    Dim wsLists As Worksheet
    Dim rngTemplates As Range
    Dim rngReinsurers As Range
    Dim i As Long, j As Long
    Dim templateName As String
    Dim reinsurer As String
    Dim foundReinsurer As Boolean

    Set wsAuto = ThisWorkbook.Sheets("Auto")
    Set wsLists = ThisWorkbook.Sheets("Lists")
    Set rngTemplates = wsLists.Range("Templates")

    For i = 1 To rngTemplates.Rows.Count
        templateName = rngTemplates.Cells(i, 1).Value
        If Trim(templateName) <> "" Then
            wsAuto.Range("H3").Value = templateName
            Application.CalculateFull
            DoEvents

            Set rngReinsurers = wsLists.Range("B4:B103")
            foundReinsurer = False
            For j = 1 To rngReinsurers.Rows.Count
                reinsurer = rngReinsurers.Cells(j, 1).Value
                If Trim(reinsurer) <> "" Then
                    foundReinsurer = True
                    wsAuto.Range("B2").Value = reinsurer
                    Application.CalculateFull
                    DoEvents
                    ProcessSingle True
                End If
            Next j
            If Not foundReinsurer Then
                Debug.Print "Skipping template '" & templateName & "' (no reinsurers found)"
            End If
        End If
    Next i
    MsgBox "All templates and reinsurers processed."
End Sub

Sub ProcessTemplate(Optional quietMode As Boolean = True)
    Dim ws As Worksheet
    Dim rng As Range
    Dim itemList As Variant
    Dim nonEmptyItems() As String
    Dim i As Integer, itemCount As Integer
    Dim originalSheet As Worksheet

    Set originalSheet = ThisWorkbook.Sheets("Auto")
    Set ws = originalSheet

    On Error GoTo ErrorHandler

    ' --- Retrieve item list from named range "Legal_Name" ---
    On Error Resume Next
    Set rng = ThisWorkbook.Names("Legal_Name").RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "The named range 'Legal_Name' does not exist."
        Debug.Print "Named range 'Legal_Name' does not exist"
        GoTo Cleanup
    End If

    itemList = rng.Value
    If Not IsArray(itemList) Then
        MsgBox "The named range 'Legal_Name' did not return an array."
        GoTo Cleanup
    End If

    ' --- Build nonEmptyItems array ---
    itemCount = 0
    For i = 1 To UBound(itemList, 1)
        If Trim(itemList(i, 1)) <> "" Then
            itemCount = itemCount + 1
            ReDim Preserve nonEmptyItems(1 To itemCount)
            nonEmptyItems(itemCount) = itemList(i, 1)
        End If
    Next i

    If itemCount = 0 Then GoTo Cleanup

    ' --- Loop through reinsurers for this template ---
    For i = 1 To itemCount
        ThisWorkbook.Sheets("Processing").Range("O36").Value = "Processing item " & i & " of " & itemCount
        DoEvents

        ws.Range("B2").Value = nonEmptyItems(i)
        Application.CalculateFull
        DoEvents

        ProcessSingle quietMode
    Next i

    Call SortAndProtectSummary

Cleanup:
    If itemCount > 0 Then ws.Range("B2").Value = nonEmptyItems(1)
    Application.CalculateFull
    DoEvents

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in ProcessTemplate: " & Err.Description & " at line " & Erl
    Debug.Print "An error occurred in ProcessTemplate: " & Err.Description & " at line " & Erl
    GoTo Cleanup
End Sub

Sub ProcessSingle(Optional quietMode As Boolean = False)
    Dim ws As Worksheet
    Dim selectedItem As Variant
    Dim originalSheet As Worksheet

    ' State preservation
    Dim prevScreenUpdating As Boolean, prevEnableEvents As Boolean, prevDisplayAlerts As Boolean
    Dim prevCalc As XlCalculation, prevFullScreen As Boolean, prevFormulaBar As Boolean
    Dim prevStatusBar As Boolean, prevRibbonVisible As Boolean, prevGridlines As Boolean
    Dim procSheet As Worksheet

    ' Set reference to original sheet immediately for error handler protection
    Set originalSheet = ThisWorkbook.Sheets("Auto")
    Set ws = originalSheet
    Set procSheet = ThisWorkbook.Sheets("Processing")

    On Error GoTo ErrorHandler

    If quietMode Then
        prevScreenUpdating = Application.ScreenUpdating
        prevEnableEvents = Application.EnableEvents
        prevDisplayAlerts = Application.DisplayAlerts
        prevCalc = Application.Calculation
        prevFullScreen = Application.DisplayFullScreen
        prevFormulaBar = Application.DisplayFormulaBar
        prevStatusBar = Application.DisplayStatusBar
        prevRibbonVisible = Application.CommandBars("Ribbon").Visible
        prevGridlines = Application.ActiveWindow.DisplayGridlines

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.DisplayFullScreen = True
        Application.DisplayFormulaBar = False
        Application.DisplayStatusBar = False
        Application.CommandBars("Ribbon").Visible = False
        Application.ActiveWindow.DisplayGridlines = False
        procSheet.Visible = xlSheetVisible
        procSheet.Activate
    Else
        procSheet.Visible = xlSheetVisible
        procSheet.Activate
    End If
    DoEvents

    selectedItem = ws.Range("B2").Value

    ' Defensive checks
    If IsError(selectedItem) Or VarType(selectedItem) <> vbString Or Trim(selectedItem) = "" Then
        Debug.Print "ProcessSingle: No valid reinsurer in B2. Skipping."
        GoTo Cleanup
    End If
    If Trim(selectedItem) = "Not in Use" Then
        Debug.Print "ProcessSingle: Skipping reinsurer 'Not in Use'."
        GoTo Cleanup
    End If

    procSheet.Range("O36").Value = "Processing single item"
    DoEvents

    ' === State synchronization: BEGIN ===
    ws.Range("B2").Value = selectedItem
    Application.CalculateFull
    Sheet5.Range("N3:N8").Calculate
    DoEvents
    Application.Wait (Now + TimeValue("0:00:01"))
    DoEvents
    ' === State synchronization: END ===

    SaveAlteredWorkbook ws, selectedItem

    DoEvents
    Worksheets("Harvested Data").Calculate
    DoEvents

    UpdateOrCreateSummaryRow
    Call SortAndProtectSummary

Cleanup:
    If quietMode Then
        Application.ScreenUpdating = prevScreenUpdating
        Application.EnableEvents = prevEnableEvents
        Application.DisplayAlerts = prevDisplayAlerts
        Application.Calculation = prevCalc
        Application.DisplayFullScreen = prevFullScreen
        Application.DisplayFormulaBar = prevFormulaBar
        Application.DisplayStatusBar = prevStatusBar
        Application.CommandBars("Ribbon").Visible = prevRibbonVisible
        If Not originalSheet Is Nothing Then originalSheet.Activate
        Application.ActiveWindow.DisplayGridlines = prevGridlines
        ThisWorkbook.Sheets("Processing").Visible = xlSheetHidden
        Application.ScreenUpdating = True
    Else
        Application.ScreenUpdating = False
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayFullScreen = False
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = True
        Application.CommandBars("Ribbon").Visible = True
        If Not originalSheet Is Nothing Then originalSheet.Activate
        Application.ActiveWindow.DisplayGridlines = True
        ThisWorkbook.Sheets("Processing").Visible = xlSheetHidden
        Application.ScreenUpdating = True
    End If

    Debug.Print "ProcessSingle completed"
    Exit Sub

ErrorHandler:
    If quietMode Then
        Application.ScreenUpdating = prevScreenUpdating
        Application.EnableEvents = prevEnableEvents
        Application.DisplayAlerts = prevDisplayAlerts
        Application.Calculation = prevCalc
        Application.DisplayFullScreen = prevFullScreen
        Application.DisplayFormulaBar = prevFormulaBar
        Application.DisplayStatusBar = prevStatusBar
        Application.CommandBars("Ribbon").Visible = prevRibbonVisible
        If Not originalSheet Is Nothing Then originalSheet.Activate
        Application.ActiveWindow.DisplayGridlines = prevGridlines
        ThisWorkbook.Sheets("Processing").Visible = xlSheetVeryHidden
        Application.ScreenUpdating = True
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayFullScreen = False
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = True
        Application.CommandBars("Ribbon").Visible = True
        If Not originalSheet Is Nothing Then originalSheet.Activate
        Application.ActiveWindow.DisplayGridlines = True
        ThisWorkbook.Sheets("Processing").Visible = xlSheetVeryHidden
        Application.ScreenUpdating = True
    End If
    MsgBox "An error occurred in ProcessSingle: " & Err.Description & " at line " & Erl
    Debug.Print "An error occurred in ProcessSingle: " & Err.Description & " at line " & Erl
    GoTo Cleanup
End Sub
Function CleanString(inputString As String) As String
    Dim cleanedString As String
    Dim i As Integer
    Dim charCode As Integer
    
    cleanedString = Trim(inputString)
    cleanedString = Replace(cleanedString, Chr(160), " ")
    For i = 1 To Len(cleanedString)
        charCode = Asc(Mid(cleanedString, i, 1))
        If charCode < 32 Or charCode > 126 Then
            cleanedString = Replace(cleanedString, Mid(cleanedString, i, 1), "")
        End If
    Next i
    cleanedString = Replace(cleanedString, "*", "")
    cleanedString = Replace(cleanedString, "?", "")
    cleanedString = Replace(cleanedString, """", "")
    cleanedString = Replace(cleanedString, "<", "")
    cleanedString = Replace(cleanedString, ">", "")
    cleanedString = Replace(cleanedString, "|", "")
    CleanString = cleanedString
End Function

Sub CreateDirectoryPath(path As String)
    Dim fso As Object
    Dim parentPath As String
    Dim pathParts() As String
    Dim i As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    pathParts = Split(path, "\")
    If Left(path, 2) = "\\" Then
        parentPath = "\\" & pathParts(2) & "\" & pathParts(3) & "\"
        i = 4
    Else
        parentPath = pathParts(0) & "\"
        i = 1
    End If
    For i = i To UBound(pathParts)
        parentPath = parentPath & pathParts(i) & "\"
        Debug.Print "Checking path: " & parentPath
        If Not fso.FolderExists(parentPath) Then
            Debug.Print "Creating folder: " & parentPath
            fso.CreateFolder parentPath
        End If
    Next i
End Sub

Function SaveAlteredWorkbook(wsAuto As Worksheet, itemName As String) As Workbook
    Dim tempWorkbook As Workbook
    Dim originalWorkbook As Workbook
    Dim directoryPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim tempFilePath As String
    Dim vbComponent As Object
    Dim invalidChars As String
    Dim i As Integer

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Set originalWorkbook = ThisWorkbook
    Debug.Print "Original workbook: " & originalWorkbook.Name

    directoryPath = CleanString(Sheet5.Range("N8").Value)
    fileName = CleanString(Sheet5.Range("N5").Value)
    Debug.Print "N8 (directory path): " & directoryPath & " (Length: " & Len(directoryPath) & ")"
    Debug.Print "N5 (file name): " & fileName & " (Length: " & Len(fileName) & ")"

    Debug.Print "Ensuring directory path exists..."
    CreateDirectoryPath directoryPath
    Debug.Print "Directory path ensured."

    fullPath = directoryPath & "\" & fileName
    tempFilePath = directoryPath & "\temp_" & fileName
    Debug.Print "Constructed file path: " & fullPath & " (Length: " & Len(fullPath) & ")"
    Debug.Print "Constructed temporary file path: " & tempFilePath & " (Length: " & Len(tempFilePath) & ")"

    Debug.Print "Checking for hidden characters..."
    Call CheckPathCharacters(fullPath)
    Debug.Print "Hidden characters check completed."

    If Len(fullPath) > 255 Then
        MsgBox "The file path is too long: " & fullPath
        Debug.Print "The file path is too long: " & fullPath
        GoTo Cleanup
    End If

    invalidChars = "<>""/|?*"
    For i = 1 To Len(invalidChars)
        If InStr(fullPath, Mid(invalidChars, i, 1)) > 0 Then
            MsgBox "The file path contains invalid characters: " & fullPath
            Debug.Print "The file path contains invalid characters: " & fullPath
            GoTo Cleanup
        End If
    Next i

    If Dir(fullPath) <> "" Then
        Debug.Print "Deleting existing file: " & fullPath
        Kill fullPath
        Debug.Print "Existing file deleted."
    End If

    On Error Resume Next
    Debug.Print "Saving original workbook as a copy..."
    originalWorkbook.SaveCopyAs tempFilePath
    If Err.Number <> 0 Then
        MsgBox "Error saving the workbook as a copy: " & Err.Description
        Debug.Print "Error saving the workbook as a copy: " & Err.Description
        Err.Clear
        GoTo Cleanup
    End If
    On Error GoTo 0
    Debug.Print "Saved workbook as a copy with temporary name."

    Debug.Print "Opening the copied workbook..."
    Set tempWorkbook = Workbooks.Open(tempFilePath)
    Debug.Print "Opened temporary workbook: " & tempWorkbook.Name

    On Error Resume Next
    Debug.Print "Hiding the Processing sheet..."
    tempWorkbook.Sheets("Processing").Visible = xlSheetVeryHidden
    On Error GoTo 0
    Debug.Print "Processing sheet hidden."

    Debug.Print "Unprotecting sheets in the temporary workbook..."
    UnprotectSheets tempWorkbook
    Debug.Print "Sheets unprotected."

    Debug.Print "Performing copy-paste operations..."
    PerformCopyPaste tempWorkbook
    Debug.Print "Copy-paste operations performed."

    Debug.Print "Breaking links..."
    BreakLinks tempWorkbook
    Debug.Print "Links broken."

    Debug.Print "Protecting sheets in the temporary workbook..."
    ProtectSheets tempWorkbook
    Debug.Print "Sheets protected."

    On Error Resume Next
    Debug.Print "Removing Module2 from the copied workbook..."
    Set vbComponent = tempWorkbook.VBProject.VBComponents("Module2")
    tempWorkbook.VBProject.VBComponents.Remove vbComponent
    On Error GoTo 0
    Debug.Print "Module2 removed from the copied workbook."

    Application.DisplayAlerts = False
    On Error Resume Next
    Debug.Print "Saving the temporary workbook with the correct file name and format..."
    tempWorkbook.SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled, ConflictResolution:=xlLocalSessionChanges
    If Err.Number <> 0 Then
        MsgBox "Error saving the workbook in SaveAlteredWorkbook: " & Err.Description
        Debug.Print "Error saving the workbook in SaveAlteredWorkbook: " & Err.Description
        Err.Clear
        Application.DisplayAlerts = True
        GoTo Cleanup
    End If
    On Error GoTo 0
    Debug.Print "Saved workbook as: " & fullPath

    Application.DisplayAlerts = True

    Debug.Print "Closing the temporary workbook..."
    tempWorkbook.Close SaveChanges:=False
    Debug.Print "Closed temporary workbook."

    If Dir(tempFilePath) <> "" Then
        Debug.Print "Deleting temporary file: " & tempFilePath
        Kill tempFilePath
        Debug.Print "Deleted temporary file: " & tempFilePath
    End If

    Debug.Print "Reactivating the original workbook..."
    originalWorkbook.Activate
    Debug.Print "Reactivated original workbook: " & originalWorkbook.Name

    Application.ScreenUpdating = True
    Set SaveAlteredWorkbook = tempWorkbook
    Exit Function

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set SaveAlteredWorkbook = Nothing
    Exit Function

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "An error occurred in SaveAlteredWorkbook: " & Err.Description
    Debug.Print "An error occurred in SaveAlteredWorkbook: " & Err.Description
    Set SaveAlteredWorkbook = Nothing
End Function


Sub CheckPathCharacters(path As String)
    Dim i As Integer
    For i = 1 To Len(path)
        Debug.Print Mid(path, i, 1) & " - " & Asc(Mid(path, i, 1))
    Next i
End Sub

Sub PerformCopyPaste(wb As Workbook)
    Dim wsDataInput As Worksheet
    Dim wsHarvestedData As Worksheet
    Dim rangesToCopy As Variant
    Dim i As Integer
    Dim sourceRange As Range
    Dim destRange As Range

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    On Error Resume Next
    Set wsDataInput = wb.Sheets("Data Input")
    Set wsHarvestedData = wb.Sheets("Harvested Data")
    On Error GoTo 0

    If wsDataInput Is Nothing Or wsHarvestedData Is Nothing Then
        MsgBox "One or both of the sheets do not exist in the workbook."
        Exit Sub
    End If

    rangesToCopy = Array( _
        Array(wsDataInput, "C1:D11"), _
        Array(wsDataInput, "G1:H11"), _
        Array(wsDataInput, "K1:L11"), _
        Array(wsDataInput, "O1:P11"), _
        Array(wsHarvestedData, "D1:Z22") _
    )

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    For i = LBound(rangesToCopy) To UBound(rangesToCopy)
        Set sourceRange = rangesToCopy(i)(0).Range(rangesToCopy(i)(1))
        Set destRange = rangesToCopy(i)(0).Range(rangesToCopy(i)(1))
        If sourceRange.MergeCells Then
            sourceRange.UnMerge
        End If
        If sourceRange.Rows.Count = destRange.Rows.Count And sourceRange.Columns.Count = destRange.Columns.Count Then
            sourceRange.Copy
            On Error Resume Next
            destRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            If Err.Number <> 0 Then
                MsgBox "Error pasting values: " & Err.Description
                Debug.Print "Error pasting values: " & Err.Description
                Err.Clear
                Exit Sub
            End If
            On Error GoTo 0
            Application.CutCopyMode = False
            DoEvents
            If Not sourceRange Is Nothing Then
                If Not sourceRange.Worksheet Is Nothing Then
                    sourceRange.Worksheet.Activate
                    sourceRange.Worksheet.Range("A1").Select
                End If
            End If
        Else
            MsgBox "Source and destination ranges are not the same size: " & rangesToCopy(i)(1)
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

    Debug.Print "Sheet4 visibility before hiding: " & Sheet4.Visible
    wb.Sheets(Sheet4.Name).Visible = xlSheetHidden
    wsDataInput.Select
    wsDataInput.Range("A1").Select
    Debug.Print "Sheet4 visibility after hiding: " & Sheet4.Visible
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred in PerformCopyPaste: " & Err.Description
End Sub

Sub UnprotectSheets(wb As Workbook, Optional password As String = "CPS@")
    Dim ws As Worksheet
    password = "CPS@"
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    For Each ws In wb.Worksheets
        ws.Unprotect password:=password
    Next ws
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred in UnprotectSheets: " & Err.Description
End Sub

Sub ProtectSheets(wb As Workbook, Optional password As String = "CPS@")
    Dim ws As Worksheet
    password = "CPS@"
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    For Each ws In wb.Worksheets
        ws.Protect password:=password
    Next ws
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred in ProtectSheets: " & Err.Description
End Sub

Sub BreakLinks(wb As Workbook, Optional password As String = "CPS@")
    Dim ext_link As Variant
    Dim linkTypes As Variant
    Dim i As Integer
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    linkTypes = Array(xlExcelLinks, xlOLELinks)
    For i = LBound(linkTypes) To UBound(linkTypes)
        If Not IsEmpty(wb.LinkSources(linkTypes(i))) Then
            For Each ext_link In wb.LinkSources(linkTypes(i))
                wb.BreakLink ext_link, linkTypes(i)
            Next ext_link
        End If
    Next i
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred in BreakLinks: " & Err.Description
End Sub


Sub CreateButtons()
    Dim ws As Worksheet
    Dim btn As OLEObject
    Dim codeModule As Object
    Dim lineNum As Long
    Dim code As String
    Dim password As String

    Set ws = ThisWorkbook.Sheets("Auto")
    password = "CPS@"

    On Error Resume Next
    ws.Unprotect password:=password
    On Error GoTo 0

    ' Delete existing buttons in K20, K24, K28
    For Each btn In ws.OLEObjects
        If btn.TopLeftCell.Address = ws.Range("K20").Address _
           Or btn.TopLeftCell.Address = ws.Range("K24").Address _
           Or btn.TopLeftCell.Address = ws.Range("K28").Address Then
            btn.Delete
        End If
    Next btn

    ' --- Create ProcessSingle button ---
    Set btn = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                Left:=ws.Range("K20").Left, _
                                Top:=ws.Range("K20").Top, _
                                Width:=ws.Range("K20").Width * 1.5, _
                                Height:=ws.Range("K20").Height * 2.5)
    With btn.Object
        .Caption = "Process Single"
    End With
    btn.Name = "btnProcessSingle"

    ' --- Create ProcessTemplate button ---
    Set btn = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                Left:=ws.Range("K24").Left, _
                                Top:=ws.Range("K24").Top, _
                                Width:=ws.Range("K24").Width * 1.5, _
                                Height:=ws.Range("K24").Height * 2.5)
    With btn.Object
        .Caption = "Process Template"
    End With
    btn.Name = "btnProcessTemplate"

    ' --- Create ProcessAllTemplates button ---
    Set btn = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                Left:=ws.Range("K28").Left, _
                                Top:=ws.Range("K28").Top, _
                                Width:=ws.Range("K28").Width * 1.5, _
                                Height:=ws.Range("K28").Height * 2.5)
    With btn.Object
        .Caption = "Process All Templates"
    End With
    btn.Name = "btnProcessAllTemplates"

    Set codeModule = ThisWorkbook.VBProject.VBComponents(ws.CodeName).codeModule

    ' Insert event handler for ProcessSingle
    If Not EventHandlerExists(codeModule, "btnProcessSingle_Click") Then
        lineNum = codeModule.CountOfLines + 1
        code = "Private Sub btnProcessSingle_Click()" & vbCrLf & _
               "    ProcessSingle" & vbCrLf & _
               "End Sub"
        codeModule.InsertLines lineNum, code
    End If

    ' Insert event handler for ProcessTemplate
    If Not EventHandlerExists(codeModule, "btnProcessTemplate_Click") Then
        lineNum = codeModule.CountOfLines + 1
        code = "Private Sub btnProcessTemplate_Click()" & vbCrLf & _
               "    ProcessTemplate" & vbCrLf & _
               "End Sub"
        codeModule.InsertLines lineNum, code
    End If

    ' Insert event handler for ProcessAllTemplates
    If Not EventHandlerExists(codeModule, "btnProcessAllTemplates_Click") Then
        lineNum = codeModule.CountOfLines + 1
        code = "Private Sub btnProcessAllTemplates_Click()" & vbCrLf & _
               "    ProcessAllTemplates" & vbCrLf & _
               "End Sub"
        codeModule.InsertLines lineNum, code
    End If

    ws.Protect password:=password
End Sub

Function EventHandlerExists(codeModule As Object, eventName As String) As Boolean
    Dim lineNum As Long
    Dim lineText As String
    EventHandlerExists = False
    For lineNum = 1 To codeModule.CountOfLines
        lineText = codeModule.Lines(lineNum, 1)
        If InStr(lineText, "Private Sub " & eventName & "()") > 0 Then
            EventHandlerExists = True
            Exit Function
        End If
    Next lineNum
End Function
Sub UpdateOrCreateSummaryRow()
    Dim summaryPath As String, summaryFileName As String, summaryFullPath As String
    Dim wbSummary As Workbook, wsSummary As Worksheet, wsLegend As Worksheet
    Dim parsedReinsurerName As String, templateType As String
    Dim nextRow As Long, foundRow As Long
    Dim alreadyOpen As Boolean
    Dim cell As Range, wb As Workbook, wsAll As Worksheet
    Dim wsAllName As String, colLetter As Variant

    Debug.Print "---- UpdateOrCreateSummaryRow START ----"

    summaryPath = ThisWorkbook.Worksheets("RAC").Range("N18").Value
    summaryFileName = ThisWorkbook.Worksheets("RAC").Range("N19").Value
    summaryFullPath = summaryPath & "\" & summaryFileName

    parsedReinsurerName = ThisWorkbook.Worksheets("Harvested Data").Range("B5").Value

    Dim cloudAdminName As String
    cloudAdminName = ThisWorkbook.Worksheets("Auto").Range("B2").Value

    templateType = ThisWorkbook.Worksheets("Auto").Range("H3").Value

    Debug.Print "parsedReinsurerName (for output): [" & parsedReinsurerName & "]"
    Debug.Print "cloudAdminName (for lookup): [" & cloudAdminName & "]"
    Debug.Print "templateType: " & templateType

    If templateType = "Test" Or templateType = "Not Reported" Then
        Debug.Print "Exiting: Template type is '" & templateType & "'"
        Exit Sub
    End If

    Debug.Print "Ensuring directory exists..."
    Call CreateDirectoryPath(summaryPath)
    Debug.Print "Directory check done."

    wsAllName = ThisWorkbook.Worksheets("Harvested Data").Range("D5").Value
    Debug.Print "wsAllName from Harvested Data!D5: '" & wsAllName & "'"
    If wsAllName = "" Then
        MsgBox "The current 'All' worksheet name is blank. Please check cell Harvested Data!D5.", vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    Set wsAll = ThisWorkbook.Worksheets(wsAllName)
    On Error GoTo 0
    If wsAll Is Nothing Then
        MsgBox "Worksheet '" & wsAllName & "' does not exist!", vbExclamation
        Exit Sub
    End If

    Dim wsLists As Worksheet
    Dim employeeName As String
    Dim listsLastRow As Long
    Dim i As Long

    Set wsLists = ThisWorkbook.Worksheets("Lists")
    listsLastRow = wsLists.Cells(wsLists.Rows.Count, "B").End(xlUp).Row
    employeeName = ""

    For i = 4 To listsLastRow
        If Trim(wsLists.Cells(i, "B").Value) = Trim(cloudAdminName) Then
            employeeName = wsLists.Cells(i, "A").Value
            Debug.Print "Employee found: [" & employeeName & "] in row " & i
            Exit For
        End If
    Next i

    If employeeName = "" Then
        Debug.Print "No employee found for reinsurer: [" & cloudAdminName & "]"
    End If

    alreadyOpen = False
    Set wbSummary = Nothing
    For Each wb In Workbooks
        If StrComp(wb.FullName, summaryFullPath, vbTextCompare) = 0 Then
            Set wbSummary = wb
            alreadyOpen = True
            Exit For
        End If
    Next wb

    If wbSummary Is Nothing Then
        On Error Resume Next
        Set wbSummary = Workbooks.Open(summaryFullPath)
        On Error GoTo 0
    End If

    If wbSummary Is Nothing Then
        Set wbSummary = Workbooks.Add
        Set wsSummary = wbSummary.Sheets(1)
        wsSummary.Name = "Summary"
        wsSummary.Range("A1").Value = "Monthly Cession File " & Format(ThisWorkbook.Worksheets("Auto").Range("H2").Value, "yyyy.mm")
        wsSummary.Range("D1").Value = "(Bal Sheet)"
        wsSummary.Range("E1").Value = "(Bal & P&L)"
        wsSummary.Range("F1").Value = "(Bal Sheet)"
        wsSummary.Range("G1").Value = "(Bal & P&L)"
        wsSummary.Range("I1").Value = "(P&L Side)"
        wsSummary.Range("A2").Value = "Assigned"
        wsSummary.Range("B2").Value = "Reinsurer"
        wsSummary.Range("C2").Value = "QB Total"
        wsSummary.Range("D2").Value = "Funding"
        wsSummary.Range("E2").Value = "Clip Fee"
        wsSummary.Range("F2").Value = "Pending Claims"
        wsSummary.Range("G2").Value = "Ceding Fee"
        wsSummary.Range("H2").Value = "Variance"
        wsSummary.Range("I2").Value = "Paid Claims"
        nextRow = 3
    Else
        Set wsSummary = wbSummary.Sheets(1)
        ' *** UNPROTECT SHEET BEFORE WRITING ***
        On Error Resume Next
        wsSummary.Unprotect password:="yourpassword"
        On Error GoTo 0
        foundRow = 0
        For Each cell In wsSummary.Range("B3:B" & wsSummary.Cells(wsSummary.Rows.Count, "B").End(xlUp).Row)
            If Trim(cell.Value) = Trim(parsedReinsurerName) And Trim(parsedReinsurerName) <> "" Then
                foundRow = cell.Row
                Exit For
            End If
        Next cell
        If foundRow > 0 Then
            nextRow = foundRow
        Else
            nextRow = wsSummary.Cells(wsSummary.Rows.Count, "B").End(xlUp).Row + 1
            If nextRow < 3 Then nextRow = 3
        End If
    End If

    Const legendSheetName As String = "Legend"
    Dim legendLastRow As Long
    Dim empColor As Long

    Dim colorPalette As Variant
    colorPalette = Array( _
        RGB(102, 204, 0), _
        RGB(255, 204, 0), _
        RGB(102, 153, 255), _
        RGB(255, 102, 102), _
        RGB(255, 153, 51), _
        RGB(153, 51, 255), _
        RGB(0, 204, 204), _
        RGB(255, 102, 255), _
        RGB(0, 153, 76), _
        RGB(255, 102, 0), _
        RGB(51, 102, 255), _
        RGB(204, 0, 102) _
    )

    On Error Resume Next
    Set wsLegend = wbSummary.Worksheets(legendSheetName)
    On Error GoTo 0

    If wsLegend Is Nothing Then
        Set wsLegend = wbSummary.Worksheets.Add(After:=wbSummary.Sheets(wbSummary.Sheets.Count))
        wsLegend.Name = legendSheetName
        wsLegend.Visible = xlSheetVisible
        wsLegend.Range("A1").Value = "Employee"
        wsLegend.Range("B1").Value = "ColorIdx"
        wsLegend.Range("C1").Value = "Sample"
    End If

    With wsLegend
        .Columns("A").ColumnWidth = 22
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 12
        With .Range("A1:C1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("A:A").HorizontalAlignment = xlCenter
        .Range("B:B").HorizontalAlignment = xlCenter
        .Range("C:C").HorizontalAlignment = xlCenter
    End With

    legendLastRow = wsLegend.Cells(wsLegend.Rows.Count, "A").End(xlUp).Row
    Dim foundLegendRow As Long
    foundLegendRow = 0

    For i = 2 To legendLastRow
        If wsLegend.Cells(i, "A").Value = employeeName Then
            foundLegendRow = i
            Exit For
        End If
    Next i

    If foundLegendRow = 0 Then
        legendLastRow = legendLastRow + 1
        wsLegend.Cells(legendLastRow, "A").Value = employeeName
        wsLegend.Cells(legendLastRow, "B").Value = (legendLastRow - 2) Mod (UBound(colorPalette) + 1)
        foundLegendRow = legendLastRow
    End If

    wsLegend.Cells(foundLegendRow, "C").Interior.Color = colorPalette(wsLegend.Cells(foundLegendRow, "B").Value)
    wsLegend.Cells(foundLegendRow, "C").Value = "" ' Just a sample cell for color
    wsLegend.Range("A" & foundLegendRow & ":C" & foundLegendRow).HorizontalAlignment = xlCenter
    empColor = colorPalette(wsLegend.Cells(foundLegendRow, "B").Value)

    wsSummary.Range("A" & nextRow & ":B" & nextRow).Interior.Color = empColor
    wsSummary.Range("D" & nextRow & ":I" & nextRow).Interior.Color = empColor
    wsSummary.Range("C" & nextRow).Interior.ColorIndex = xlNone

    With wsSummary.Range("A" & nextRow)
        .Font.Name = "Aptos"
        .Font.Size = 11
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
    End With

    wsSummary.Range("B" & nextRow).Value = parsedReinsurerName
    wsSummary.Range("C" & nextRow).Value = ""
    wsSummary.Range("D" & nextRow).Value = wsAll.Range("F20").Value
    wsSummary.Range("E" & nextRow).Value = wsAll.Range("F14").Value
    wsSummary.Range("F" & nextRow).Value = wsAll.Range("F16").Value
    wsSummary.Range("G" & nextRow).Value = wsAll.Range("F15").Value
    wsSummary.Range("H" & nextRow).Formula = "=IF(C" & nextRow & "="""","""",(G" & nextRow & "+F" & nextRow & "+E" & nextRow & "+D" & nextRow & ")-C" & nextRow & ")"
    wsSummary.Range("I" & nextRow).Value = wsAll.Range("F17").Value

    With wsSummary
        .Columns("A").ColumnWidth = 22
        .Columns("B").ColumnWidth = 60
        .Columns("C:I").ColumnWidth = 15
        With .Range("A1")
            .Font.Name = "Aptos"
            .Font.Size = 11
            .Font.Bold = True
            .HorizontalAlignment = xlLeft
        End With
        For Each colLetter In Array("D", "E", "F", "G", "I")
            With .Range(colLetter & "1")
                .Font.Name = "Aptos"
                .Font.Size = 8
                .Font.Italic = True
                .HorizontalAlignment = xlCenter
            End With
        Next colLetter
        With .Range("A2")
            .Font.Name = "Aptos"
            .Font.Size = 11
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        For Each colLetter In Array("B", "C", "D", "E", "F", "G", "H", "I")
            With .Range(colLetter & "2")
                .Font.Name = "Aptos"
                .Font.Size = 11
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
        Next colLetter
        With .Range("B" & nextRow)
            .Font.Name = "Aptos"
            .Font.Size = 11
            .Font.Bold = False
            .HorizontalAlignment = xlLeft
            .NumberFormat = "$#,##0.00"
        End With
        For Each colLetter In Array("C", "D", "E", "F", "G", "H", "I")
            With .Range(colLetter & nextRow)
                .Font.Name = "Aptos"
                .Font.Size = 11
                .Font.Bold = False
                .HorizontalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
            End With
        Next colLetter
    End With

    Application.DisplayAlerts = False
    wbSummary.SaveAs summaryFullPath
    If Not alreadyOpen Then
        wbSummary.Close SaveChanges:=False
    End If
    Application.DisplayAlerts = True
    Debug.Print "---- UpdateOrCreateSummaryRow END ----"
End Sub
Sub SortAndProtectSummary()
    Dim summaryPath As String, summaryFileName As String, summaryFullPath As String
    Dim wbSummary As Workbook, wsSummary As Worksheet
    Dim lastRow As Long

    summaryPath = ThisWorkbook.Worksheets("RAC").Range("N18").Value
    summaryFileName = ThisWorkbook.Worksheets("RAC").Range("N19").Value
    summaryFullPath = summaryPath & "\" & summaryFileName

    Set wbSummary = Workbooks.Open(summaryFullPath)
    Set wsSummary = wbSummary.Sheets(1) ' Adjust if needed

    lastRow = wsSummary.Cells(wsSummary.Rows.Count, "A").End(xlUp).Row

    ' Sort by Employee Name
    With wsSummary.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsSummary.Range("A3:A" & lastRow), Order:=xlAscending
        .SetRange wsSummary.Range("A2:I" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' Unlock Col C for data entry, lock others
    wsSummary.Range("C3:C" & lastRow).Locked = False
    wsSummary.Range("A3:B" & lastRow).Locked = True
    wsSummary.Range("D3:I" & lastRow).Locked = True

    wsSummary.Protect password:="yourpassword", UserInterfaceOnly:=True

    wbSummary.Save
    wbSummary.Close SaveChanges:=False
End Sub




