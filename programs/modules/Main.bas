Attribute VB_Name = "Main"
Option Explicit
' tool -> references -> Microsoft Scripting Runtime

Public Sub ProcessInputFiles()
    Dim inputSheetValue() As String
    Dim targetSheetValue() As String
    Dim outputValue As Variant
    Dim templateFilePath As String
    templateFilePath = GetSingleExcelFile("input\before")
    If templateFilePath = Empty Then
        Exit Sub
    End If
    Dim targetFilePath As String
    targetFilePath = GetSingleExcelFile("input\after")
    If targetFilePath = Empty Then
        Exit Sub
    End If
    Dim workingFolderPath As String
    workingFolderPath = GetTargetPath(GetParentPath(), "temp")
    Call GetOrCreateFolder(workingFolderPath)
    OUTPUT_FILENAME = Mid(targetFilePath, InStrRev(targetFilePath, "\") + 1)
    FileCopy templateFilePath, workingFolderPath & "\before.xlsx"
    FileCopy targetFilePath, workingFolderPath & "\after.xlsx"
On Error GoTo FINL_L
    Dim templateFile As Workbook
    Dim targetFile As Workbook
    Set templateFile = Workbooks.Open(workingFolderPath & "\before.xlsx")
    Set targetFile = Workbooks.Open(workingFolderPath & "\after.xlsx")
    
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
FINL_L:
    Call CloseWorkbookWithoutSaving(templateFile)
    Call CloseWorkbookWithoutSaving(targetFile)
    
    ThisWorkbook.Saved = True
End Sub

Private Function CompareArrays(inputSheetValue As Variant, targetSheetValue As Variant) As Variant
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
    
    rowsCount = UBound(targetSheetValue, 1) - LBound(targetSheetValue, 1) + 1
    colsCount = UBound(targetSheetValue, 2) - LBound(targetSheetValue, 2) + 1
    ReDim resultArray(0 To rowsCount - 1, 0 To colsCount - 1)
    lastRow = rowsCount
    addSeq = lastRow
    For i = LBound(targetSheetValue) To UBound(targetSheetValue)
        found = False
        For j = 0 To colsCount - 1
            resultArray(i, j) = targetSheetValue(i, j)
        Next j
        For j = LBound(inputSheetValue) To UBound(inputSheetValue)
            If (resultArray(i, keyIndex("phoneNumber")) = inputSheetValue(j, keyIndex("phoneNumber"))) Or _
               (resultArray(i, keyIndex("facilityName")) = inputSheetValue(j, keyIndex("facilityName"))) Or _
               (resultArray(i, keyIndex("facilityAddress")) = inputSheetValue(j, keyIndex("facilityAddress"))) Then
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
    
    For i = LBound(inputSheetValue) To UBound(inputSheetValue)
        found = False
        For j = LBound(resultArray) To UBound(resultArray)
            If (inputSheetValue(i, keyIndex("phoneNumber")) = resultArray(j, keyIndex("phoneNumber"))) Or _
               (inputSheetValue(i, keyIndex("facilityName")) = resultArray(j, keyIndex("facilityName"))) Or _
               (inputSheetValue(i, keyIndex("facilityAddress")) = resultArray(j, keyIndex("facilityAddress"))) Then
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
    Dim outputSheet As Worksheet
    Dim addRowCount As Integer
    Dim addStartRow As Integer
    Dim addEndRow As Integer
    Dim outputLastRangeAddress As String
    
    outputFolder = GetTargetPath(GetParentPath(), outputFolderName)
    Call GetOrCreateFolder(outputFolder)
    outputFilename = outputFolder & "\" & OUTPUT_FILENAME
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



