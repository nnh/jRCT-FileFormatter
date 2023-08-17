Attribute VB_Name = "Main"
Option Explicit
' tool -> references -> Microsoft Scripting Runtime

Public Sub ProcessInputFiles()
    Dim inputSheetValue() As String
    Dim targetSheetValue() As String
    Dim outputValue As Variant
    Dim templateFile As Workbook
    Set templateFile = OpenWorkbook("input", INPUT_WORKBOOK_NAME)
    Dim targetFile As Workbook
    Set targetFile = OpenWorkbook("input", TARGET_WORKBOOK_NAME)
    
    inputSheetValue = GetInputSheetValuesToArray(templateFile.Worksheets(DEFAULT_SHEET_NAME))
    targetSheetValue = GetTargetSheetValuesToArray(targetFile.Worksheets(DEFAULT_SHEET_NAME))
    outputValue = CompareArrays(inputSheetValue, targetSheetValue)
    Dim i As Integer
    Dim keys As Dictionary
    Set keys = CreateAssociativeArrayKeyIndex()
    Dim tempSeq As String

    For i = LBound(outputValue) To UBound(outputValue)
        tempSeq = outputValue(i, keys("seqNo") + 1)
        If tempSeq <> Empty And IsNumeric(tempSeq) Then
            outputValue(i, keys("seqNo") + 1) = i - 1
        End If
    Next i
    
    Call ExportToFile(templateFile, outputValue)
    Call CloseWorkbookWithoutSaving(templateFile)
    Call CloseWorkbookWithoutSaving(targetFile)
    
    ThisWorkbook.Saved = True
End Sub

Private Function CompareArrays(inputSheetValue As Variant, targetSheetValue As Variant) As Variant
    Dim inputKeys As Variant
    Dim targetKeys As Variant
    Dim keyIndex As Dictionary
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim found As Boolean
    Dim rowsCount As Integer
    Dim colsCount As Integer
    Dim resultArray() As String
    Dim lastRow As Integer
    Dim addSeq As Integer
    Dim tempWs As Worksheet
    Dim outputRange As Range
    Dim targetRangeAddress As String
    Dim deleteRowValues() As String
    ReDim deleteRowValues(TARGET_LAST_COLUMN)
    deleteRowValues = GetDeleteRowValues()
    
    Set keyIndex = CreateAssociativeArrayKeyIndex()
    inputKeys = GenerateKey(inputSheetValue)
    targetKeys = GenerateKey(targetSheetValue)
    
    rowsCount = UBound(targetSheetValue, 1) - LBound(targetSheetValue, 1) + 1
    colsCount = UBound(targetSheetValue, 2) - LBound(targetSheetValue, 2) + 1
    ReDim resultArray(0 To rowsCount - 1, 0 To colsCount - 1)
    lastRow = rowsCount
    addSeq = lastRow
    
    For i = LBound(targetKeys) To UBound(targetKeys)
        found = False
        For j = 0 To colsCount - 1
            resultArray(i, j) = targetSheetValue(i, j)
        Next j
        For j = LBound(inputKeys) To UBound(inputKeys)
            If targetKeys(i) = inputKeys(j) Then
                resultArray(i, keyIndex("seqNo")) = inputSheetValue(j, keyIndex("seqNo"))
                found = True
                Exit For
            End If
        Next j
        
        If Not found Then
            addSeq = addSeq + 1
            If debugFlag Then
                resultArray(i, 0) = "add"
            End If
            resultArray(i, keyIndex("seqNo")) = addSeq
        End If
    Next i
    
    Set tempWs = GetOrCreateSheet("temp")
    tempWs.Cells.Clear
    Set outputRange = tempWs.Cells(1, 1).Resize(rowsCount, colsCount)
    outputRange.Value = resultArray
    
    For i = LBound(inputKeys) To UBound(inputKeys)
        found = False
        For j = LBound(targetKeys) To UBound(targetKeys)
            If inputKeys(i) = targetKeys(j) Then
                found = True
                Exit For
            End If
        Next j
        
        If Not found Then
            lastRow = lastRow + 1
            For k = 1 To colsCount - 1
                If deleteRowValues(k) <> "" Then
                    tempWs.Cells(lastRow, k + 1).Value = deleteRowValues(k)
                Else
                    tempWs.Cells(lastRow, k + 1).Value = inputSheetValue(i, k)
                End If
            Next k
            If debugFlag Then
                tempWs.Cells(lastRow, 1).Value = "del"
            End If
        End If
    Next i
    targetRangeAddress = SortOutputValues(tempWs, lastRow, colsCount)
    CompareArrays = tempWs.Range(targetRangeAddress).Value
End Function

Private Sub ExportToFile(templateFile As Workbook, setValues As Variant)
    Const outputFolderName As String = "output"
    Const inputStartRow As Integer = 2
    Dim outputFolder As String
    Dim outputFilename As String
    Dim fso As Scripting.FileSystemObject
    Dim outputSheet As Worksheet
    Dim addRowCount As Integer
    Dim addStartRow As Integer
    Dim addEndRow As Integer
    Dim outputLastRangeAddress As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    outputFolder = GetTargetPath(GetParentPath(), outputFolderName)
    If Not fso.FolderExists(outputFolder) Then
        fso.CreateFolder outputFolder
    End If
    outputFilename = outputFolder & "\" & templateFile.Name
    Set outputSheet = templateFile.Worksheets(DEFAULT_SHEET_NAME)
    addRowCount = UBound(setValues) - (INPUT_LAST_ROW - COLUMNNAME_ROW)
    addStartRow = COLUMNNAME_ROW + 1
    addEndRow = COLUMNNAME_ROW + addRowCount
    With outputSheet
        .Rows(addStartRow & ":" & addEndRow).Insert shift:=xlDown
        outputLastRangeAddress = .Cells(addStartRow + UBound(setValues) - 1, INPUT_LAST_COLUMN).Address
        .Range("A" & addStartRow & ":" & outputLastRangeAddress).Value = setValues
        .Rows(addStartRow).Delete
        .Columns("AD:AD").NumberFormatLocal = "yyyy-mm-dd;@"
        .Rows(addStartRow & ":" & addEndRow).Font.Bold = False
        .Columns("B:B").Font.Bold = True
    End With

    Call SaveWorkbook(templateFile, outputFilename)
End Sub




