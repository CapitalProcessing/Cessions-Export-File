Attribute VB_Name = "Module2"
Sub ProcessTemplate()
    Dim ws As Worksheet
    Dim itemList As Variant
    Dim nonEmptyItems() As String
    Dim i As Integer, itemCount As Integer
    Dim originalSheet As Worksheet
    Dim rng As Range

    On Error GoTo ErrorHandler
    Debug.Print "Starting ProcessTemplate"

    ' Unhide and switch to the "Processing" sheet
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        ThisWorkbook.Sheets("Processing").Visible = xlSheetVisible
        ThisWorkbook.Sheets("Processing").Activate
        .DisplayFullScreen = True
        .DisplayFormulaBar = False
        .DisplayStatusBar = False
        .CommandBars("Ribbon").Visible = False
        .ActiveWindow.DisplayGridlines = False
    End With
    DoEvents
    Debug.Print "Switched to Processing sheet"

    ' Store the original sheet
    Set originalSheet = ThisWorkbook.Sheets("Auto")
    Debug.Print "Original sheet set"

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Auto")
    Debug.Print "Worksheet set to Auto"

    ' Check if the named range "Legal_Name" exists
    On Error Resume Next
    Set rng = ThisWorkbook.Names("Legal_Name").RefersToRange
    If rng Is Nothing Then
        MsgBox "The named range 'Legal_Name' does not exist."
        Debug.Print "Named range 'Legal_Name' does not exist"
        GoTo Cleanup
    End If
    On Error GoTo 0
    Debug.Print "Named range 'Legal_Name' exists"

    ' Convert the range to an array
    itemList = rng.Value
    Debug.Print "Items retrieved from named range 'Legal_Name'"

    ' Debugging: Print the type and value of itemList
    Debug.Print "Type of itemList: " & TypeName(itemList)
    If IsArray(itemList) Then
        Debug.Print "itemList is an array with " & UBound(itemList, 1) & " rows."
    Else
        Debug.Print "itemList is not an array."
        MsgBox "The named range 'Legal_Name' did not return an array."
        GoTo Cleanup
    End If

    ' Create an array with non-empty items only
    itemCount = 0
    For i = 1 To UBound(itemList, 1)
        If itemList(i, 1) <> "" Then
            itemCount = itemCount + 1
            ReDim Preserve nonEmptyItems(1 To itemCount)
            nonEmptyItems(itemCount) = itemList(i, 1)
        End If
    Next i
    Debug.Print "Non-empty items array created"

    ' Debugging: Print the nonEmptyItems array
    For i = 1 To itemCount
        Debug.Print "nonEmptyItems(" & i & "): " & nonEmptyItems(i)
    Next i

    ' Exit if there are no items
    If itemCount = 0 Then
        GoTo Cleanup
    End If

    ' Call ProcessItems to handle item-wise processing
    ProcessItems ws, nonEmptyItems

Cleanup:
    ' Restore the original view settings and switch back to the original sheet
    With Application
        .ScreenUpdating = False
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayFullScreen = False
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
        .CommandBars("Ribbon").Visible = True
        originalSheet.Activate
        .ActiveWindow.DisplayGridlines = True
        ThisWorkbook.Sheets("Processing").Visible = xlSheetHidden
        .ScreenUpdating = True
    End With
    Debug.Print "ProcessTemplate completed"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    With Application
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayFullScreen = False
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
        .CommandBars("Ribbon").Visible = True
        originalSheet.Activate
        .ActiveWindow.DisplayGridlines = True
        ThisWorkbook.Sheets("Processing").Visible = xlSheetVeryHidden
    End With
    MsgBox "An error occurred in ProcessTemplate: " & Err.Description & " at line " & Erl
    Debug.Print "An error occurred in ProcessTemplate: " & Err.Description & " at line " & Erl
End Sub

Sub ProcessItems(ws As Worksheet, nonEmptyItems() As String)
    Dim i As Integer, itemCount As Integer
    Dim originalSheet As Worksheet

    On Error GoTo ErrorHandler
    Debug.Print "Starting ProcessItems"

    ' Store the original sheet
    Set originalSheet = ThisWorkbook.Sheets("Auto")
    itemCount = UBound(nonEmptyItems)

    ' Loop through each non-empty item
    For i = 1 To itemCount
        ' Update the progress indicator
        ThisWorkbook.Sheets("Processing").Range("O36").Value = "Processing item " & i & " of " & itemCount
        DoEvents

        ' Set the drop-down list cell to the current item
        ws.Range("B2").Value = nonEmptyItems(i)

        ' Force recalculation to ensure the change is registered
        Application.CalculateFull

        ' Explicitly recalculate the specific cells in Sheet5
        Sheet5.Range("N3:N8").Calculate

        ' Allow the system to process other events
        DoEvents

        ' Add a small delay
        Application.Wait (Now + TimeValue("0:00:02"))

        ' Debugging message
        Debug.Print "Processing item: " & nonEmptyItems(i)
        Debug.Print "Value in ws.Range('B2'): " & ws.Range("B2").Value

        ' Call SaveAlteredWorkbook to handle alterations and saving
        SaveAlteredWorkbook ws, nonEmptyItems(i)
    Next i

    Debug.Print "ProcessItems completed"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    With Application
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayFullScreen = False
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
        .CommandBars("Ribbon").Visible = True
        originalSheet.Activate
        .ActiveWindow.DisplayGridlines = True
        ThisWorkbook.Sheets("Processing").Visible = xlSheetVeryHidden
    End With
    MsgBox "An error occurred in ProcessItems: " & Err.Description & " at line " & Erl
    Debug.Print "An error occurred in ProcessItems: " & Err.Description & " at line " & Erl
End Sub

' Backward compatibility wrapper - calls ProcessTemplate
Sub ProcessItemsOriginal()
    ProcessTemplate
End Sub
Function CleanString(inputString As String) As String
    Dim cleanedString As String
    Dim i As Integer
    Dim charCode As Integer
    
    cleanedString = Trim(inputString)
    cleanedString = Replace(cleanedString, Chr(160), " ") ' Replace non-breaking spaces
    
    ' Remove problematic characters
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
    
    ' Do not remove backslashes, colons, or periods
    CleanString = cleanedString
End Function
Sub CreateDirectoryPath(path As String)
    Dim fso As Object
    Dim parentPath As String
    Dim pathParts() As String
    Dim i As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Split the path into its components
    pathParts = Split(path, "\")
    
    ' Handle UNC paths (e.g., \\server\share)
    If Left(path, 2) = "\\" Then
        parentPath = "\\" & pathParts(2) & "\" & pathParts(3) & "\"
        i = 4
    Else
        parentPath = pathParts(0) & "\"
        i = 1
    End If
    
    ' Reconstruct the path and create each part if it doesn't exist
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

    ' Get the directory path from N8 and the file name from N5
    directoryPath = CleanString(Sheet5.Range("N8").Value)
    fileName = CleanString(Sheet5.Range("N5").Value)
    Debug.Print "N8 (directory path): " & directoryPath & " (Length: " & Len(directoryPath) & ")"
    Debug.Print "N5 (file name): " & fileName & " (Length: " & Len(fileName) & ")"

    ' Ensure the directory exists
    Debug.Print "Ensuring directory path exists..."
    CreateDirectoryPath directoryPath
    Debug.Print "Directory path ensured."

    ' Construct the full path with the file name
    fullPath = directoryPath & "\" & fileName
    tempFilePath = directoryPath & "\temp_" & fileName
    Debug.Print "Constructed file path: " & fullPath & " (Length: " & Len(fullPath) & ")"
    Debug.Print "Constructed temporary file path: " & tempFilePath & " (Length: " & Len(tempFilePath) & ")"

    ' Check for hidden characters
    Debug.Print "Checking for hidden characters..."
    Call CheckPathCharacters(fullPath)
    Debug.Print "Hidden characters check completed."

    ' Check if the full path is valid
    If Len(fullPath) > 255 Then
        MsgBox "The file path is too long: " & fullPath
        Debug.Print "The file path is too long: " & fullPath
        GoTo Cleanup
    End If

    ' Check for invalid characters in the file path
    invalidChars = "<>""/|?*"
    For i = 1 To Len(invalidChars)
        If InStr(fullPath, Mid(invalidChars, i, 1)) > 0 Then
            MsgBox "The file path contains invalid characters: " & fullPath
            Debug.Print "The file path contains invalid characters: " & fullPath
            GoTo Cleanup
        End If
    Next i

    ' Check if the file already exists and delete it
    If Dir(fullPath) <> "" Then
        Debug.Print "Deleting existing file: " & fullPath
        Kill fullPath
        Debug.Print "Existing file deleted."
    End If

    ' Save the original workbook as a copy with a temporary name
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

    ' Open the copied workbook
    Debug.Print "Opening the copied workbook..."
    Set tempWorkbook = Workbooks.Open(tempFilePath)
    Debug.Print "Opened temporary workbook: " & tempWorkbook.Name

    ' Perform all necessary alterations on the temporary workbook
    ' Hide the Processing sheet instead of deleting it
    On Error Resume Next
    Debug.Print "Hiding the Processing sheet..."
    tempWorkbook.Sheets("Processing").Visible = xlSheetVeryHidden
    On Error GoTo 0
    Debug.Print "Processing sheet hidden."

    ' Unprotect sheets in the temporary workbook
    Debug.Print "Unprotecting sheets in the temporary workbook..."
    UnprotectSheets tempWorkbook
    Debug.Print "Sheets unprotected."

    ' Perform copy-paste operations on the temporary workbook
    Debug.Print "Performing copy-paste operations..."
    PerformCopyPaste tempWorkbook
    Debug.Print "Copy-paste operations performed."

    ' Break links
    Debug.Print "Breaking links..."
    BreakLinks tempWorkbook
    Debug.Print "Links broken."

    ' Protect sheets in the temporary workbook
    Debug.Print "Protecting sheets in the temporary workbook..."
    ProtectSheets tempWorkbook
    Debug.Print "Sheets protected."

    ' Remove Module2 from the copied workbook
    On Error Resume Next
    Debug.Print "Removing Module2 from the copied workbook..."
    Set vbComponent = tempWorkbook.VBProject.VBComponents("Module2")
    tempWorkbook.VBProject.VBComponents.Remove vbComponent
    On Error GoTo 0
    Debug.Print "Module2 removed from the copied workbook."

    ' Disable alerts to prevent the "overwrite" prompt
    Application.DisplayAlerts = False

    ' Save the temporary workbook with the correct file name and format
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

    ' Re-enable alerts
    Application.DisplayAlerts = True

    ' Close the temporary workbook without saving changes
    Debug.Print "Closing the temporary workbook..."
    tempWorkbook.Close SaveChanges:=False
    Debug.Print "Closed temporary workbook."

    ' Delete the temporary file
    If Dir(tempFilePath) <> "" Then
        Debug.Print "Deleting temporary file: " & tempFilePath
        Kill tempFilePath
        Debug.Print "Deleted temporary file: " & tempFilePath
    End If

    ' Reactivate the original workbook
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
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Set worksheets using sheet names
    On Error Resume Next
    Set wsDataInput = wb.Sheets("Data Input") ' Changed sheet name
    Set wsHarvestedData = wb.Sheets("Harvested Data") ' Replace with the actual sheet name
    On Error GoTo 0

    If wsDataInput Is Nothing Or wsHarvestedData Is Nothing Then
        MsgBox "One or both of the sheets do not exist in the workbook."
        Exit Sub
    End If

    ' Ranges to copy with associated worksheets
    rangesToCopy = Array( _
        Array(wsDataInput, "C1:D11"), _
        Array(wsDataInput, "G1:H11"), _
        Array(wsDataInput, "K1:L11"), _
        Array(wsDataInput, "O1:P11"), _
        Array(wsHarvestedData, "D1:Z22") _
    )

    ' Optimize performance
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    ' Copy and paste values
    For i = LBound(rangesToCopy) To UBound(rangesToCopy)
        Set sourceRange = rangesToCopy(i)(0).Range(rangesToCopy(i)(1))
        Set destRange = rangesToCopy(i)(0).Range(rangesToCopy(i)(1))

        ' Debugging: Print range addresses and sizes
        Debug.Print "Source range: " & sourceRange.Address & vbCrLf & "Destination range: " & destRange.Address & vbCrLf & _
                    "Source rows: " & sourceRange.Rows.Count & ", Source columns: " & sourceRange.Columns.Count & vbCrLf & _
                    "Destination rows: " & destRange.Rows.Count & ", Destination columns: " & destRange.Columns.Count

        ' Unmerge cells in the source range
        If sourceRange.MergeCells Then
            sourceRange.UnMerge
        End If

        ' Ensure the destination range is the same size as the source range
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
            DoEvents ' Allow time for clipboard to reset

            ' Debugging: Check if sourceRange and its worksheet are valid
            If Not sourceRange Is Nothing Then
                Debug.Print "sourceRange is set: " & sourceRange.Address
                If Not sourceRange.Worksheet Is Nothing Then
                    Debug.Print "sourceRange.Worksheet is set: " & sourceRange.Worksheet.Name
                    ' Activate the worksheet before selecting A1
                    sourceRange.Worksheet.Activate
                    sourceRange.Worksheet.Range("A1").Select
                Else
                    Debug.Print "sourceRange.Worksheet is Nothing"
                End If
            Else
                Debug.Print "sourceRange is Nothing"
            End If
        Else
            MsgBox "Source and destination ranges are not the same size: " & rangesToCopy(i)(1)
        End If
    Next i

    ' Reset application settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

    ' Debugging: Print visibility status before hiding
    Debug.Print "Sheet4 visibility before hiding: " & Sheet4.Visible

    ' Hide Sheet4 in the temporary workbook
    wb.Sheets(Sheet4.Name).Visible = xlSheetHidden ' Use the codename directly

    ' Select A1 on the destination worksheet to ensure the copy and paste box is cleared
    wsDataInput.Select
    wsDataInput.Range("A1").Select

    ' Debugging: Print visibility status after hiding
    Debug.Print "Sheet4 visibility after hiding: " & Sheet4.Visible

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred in PerformCopyPaste: " & Err.Description
End Sub
Sub UnprotectSheets(wb As Workbook, Optional password As String = "CPS@")
    Dim ws As Worksheet
    password = "CPS@" ' Protection password

    On Error GoTo ErrorHandler
    ' Turn off screen updating
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
    password = "CPS@" ' Protection password
    
    On Error GoTo ErrorHandler
    ' Turn off screen updating
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
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Define the types of links to break
    linkTypes = Array(xlExcelLinks, xlOLELinks, xlPublisherLinks)

    ' Loop through each type of link
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
Sub ProcessSingle()
    Dim ws As Worksheet
    Dim selectedItem As String
    Dim originalSheet As Worksheet
    On Error GoTo ErrorHandler
    Debug.Print "Starting ProcessSingle"
    
    ' Unhide and switch to the "Processing" sheet
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        ThisWorkbook.Sheets("Processing").Visible = xlSheetVisible
        ThisWorkbook.Sheets("Processing").Activate
        .DisplayFullScreen = True
        .DisplayFormulaBar = False
        .DisplayStatusBar = False
        .CommandBars("Ribbon").Visible = False
        .ActiveWindow.DisplayGridlines = False
    End With
    DoEvents
    Debug.Print "Switched to Processing sheet"
    
    ' Store the original sheet
    Set originalSheet = ThisWorkbook.Sheets("Auto")
    Debug.Print "Original sheet set"
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Auto")
    Debug.Print "Worksheet set to Auto"
    
    ' Get the selected item from the drop-down list in B2
    selectedItem = ws.Range("B2").Value
    If selectedItem = "" Then
        MsgBox "Please select a Legal Name from the drop-down list in cell B2."
        GoTo Cleanup
    End If
    
    Debug.Print "Selected item: " & selectedItem
    
    ' Update the progress indicator
    ThisWorkbook.Sheets("Processing").Range("O36").Value = "Processing single item"
    DoEvents
    
    ' Set the drop-down list cell to the selected item
    ws.Range("B2").Value = selectedItem
    
    ' Force recalculation to ensure the change is registered
    Application.CalculateFull
    
    ' Explicitly recalculate the specific cells in Sheet5
    Sheet5.Range("N3:N8").Calculate
    
    ' Allow the system to process other events
    DoEvents
    
    ' Add a small delay
    Application.Wait (Now + TimeValue("0:00:01"))
    
    ' Debugging message
    Debug.Print "Processing item: " & selectedItem
    Debug.Print "Value in ws.Range('B2'): " & ws.Range("B2").Value
    
    ' Call SaveAlteredWorkbook to handle alterations and saving
    SaveAlteredWorkbook ws, selectedItem
    
Cleanup:
    ' Restore the original view settings and switch back to the original sheet
    With Application
        .ScreenUpdating = False
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayFullScreen = False
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
        .CommandBars("Ribbon").Visible = True
        originalSheet.Activate
        .ActiveWindow.DisplayGridlines = True
        ThisWorkbook.Sheets("Processing").Visible = xlSheetHidden
        .ScreenUpdating = True
    End With
    Debug.Print "ProcessSingle completed"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    With Application
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayFullScreen = False
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
        .CommandBars("Ribbon").Visible = True
        originalSheet.Activate
        .ActiveWindow.DisplayGridlines = True
        ThisWorkbook.Sheets("Processing").Visible = xlSheetVeryHidden
    End With
    MsgBox "An error occurred in ProcessSingle: " & Err.Description & " at line " & Erl
    Debug.Print "An error occurred in ProcessSingle: " & Err.Description & " at line " & Erl
End Sub
Sub CreateButtons()
    ' Creates two buttons on Auto and removes prior buttons
    Dim ws As Worksheet
    Dim btn As OLEObject
    Dim password As String
    Dim codeModule As Object
    Dim lineNum As Long
    Dim code As String

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Auto")
    password = "CPS@" ' Replace with your actual password if the sheet is protected

    ' Unprotect the worksheet
    On Error Resume Next
    ws.Unprotect password:=password
    On Error GoTo 0

    ' Delete any existing button at K20 and K24
    For Each btn In ws.OLEObjects
        If btn.TopLeftCell.Address = ws.Range("K20").Address Or btn.TopLeftCell.Address = ws.Range("K24").Address Then
            btn.Delete
        End If
    Next btn

    ' Add a new button at K20 for ProcessTemplate
    Set btn = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                Left:=ws.Range("K20").Left, _
                                Top:=ws.Range("K20").Top, _
                                Width:=ws.Range("K20").Width * 1.5, _
                                Height:=ws.Range("K20").Height * 2.5) ' Make the button twice as tall
    With btn.Object
        .Caption = "Run ProcessTemplate"
    End With
    btn.Name = "btnProcessTemplate"

    ' Add a new button at K24 for ProcessSingle
    Set btn = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                Left:=ws.Range("K24").Left, _
                                Top:=ws.Range("K24").Top, _
                                Width:=ws.Range("K24").Width * 1.5, _
                                Height:=ws.Range("K24").Height * 2.5) ' Make the button twice as tall
    With btn.Object
        .Caption = "Run ProcessSingle"
    End With
    btn.Name = "btnProcessSingle"

    ' Get the code module for the worksheet
    Set codeModule = ThisWorkbook.VBProject.VBComponents(ws.CodeName).codeModule

    ' Check if the event handler for btnProcessTemplate_Click already exists
    If Not EventHandlerExists(codeModule, "btnProcessTemplate_Click") Then
        ' Add the event handler for btnProcessTemplate_Click
        lineNum = codeModule.CountOfLines + 1
        code = "Private Sub btnProcessTemplate_Click()" & vbCrLf & _
               "    ProcessTemplate" & vbCrLf & _
               "End Sub"
        codeModule.InsertLines lineNum, code
    End If

    ' Check if the event handler for btnProcessSingle_Click already exists
    If Not EventHandlerExists(codeModule, "btnProcessSingle_Click") Then
        ' Add the event handler for btnProcessSingle_Click
        lineNum = codeModule.CountOfLines + 1
        code = "Private Sub btnProcessSingle_Click()" & vbCrLf & _
               "    ProcessSingle" & vbCrLf & _
               "End Sub"
        codeModule.InsertLines lineNum, code
    End If

    ' Protect the worksheet again
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
