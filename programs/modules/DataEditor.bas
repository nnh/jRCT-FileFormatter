Attribute VB_Name = "DataEditor"
Option Explicit

Public Function GetInputSheetValuesToArray(ws As Worksheet) As String()
    Dim rowAndColumnNumber As Variant
    rowAndColumnNumber = GetLastRowAndColumnNumber(ws)
    INPUT_LAST_ROW = rowAndColumnNumber(0)
    INPUT_LAST_COLUMN = rowAndColumnNumber(1)
    GetInputSheetValuesToArray = GetSheetValuesToArray(ws, INPUT_LAST_ROW, INPUT_LAST_COLUMN)
End Function

Public Function GetTargetSheetValuesToArray(ws As Worksheet) As String()
    Dim rowAndColumnNumber As Variant
    rowAndColumnNumber = GetLastRowAndColumnNumber(ws)
    TARGET_LAST_ROW = rowAndColumnNumber(0)
    TARGET_LAST_COLUMN = rowAndColumnNumber(1)
    GetTargetSheetValuesToArray = GetSheetValuesToArray(ws, TARGET_LAST_ROW, TARGET_LAST_COLUMN)
End Function

Private Function GetSheetValuesToArray(ws As Worksheet, lastRow As Integer, lastColumn As Integer) As String()
    Dim dataRange As Range
    Dim dataArray() As String
    Dim numRows As Integer
    Dim numCols As Integer
    Dim i As Integer
    Dim j As Integer
    Dim rowsCount As Integer
    Dim colsCount As Integer
    Dim count As Variant
    Dim dataValue As Variant
    
    Set dataRange = ws.Range(ws.Cells(COLUMNNAME_ROW, 1), ws.Cells(lastRow, lastColumn))
    count = GetArrayDimensions(dataRange.Value)
    
    dataValue = dataRange.Value
    rowsCount = count(0)
    colsCount = count(1)
    ReDim dataArray(0 To rowsCount - 1, 0 To colsCount - 1)
    For i = 0 To rowsCount - 1
        For j = 0 To colsCount - 1
            dataArray(i, j) = CStr(dataValue(i + 1, j + 1))
        Next j
    Next i
    GetSheetValuesToArray = dataArray
End Function

Public Function SortOutputValues(targetWorksheet As Worksheet, lastRow As Integer, lastColumn As Integer) As String
    Dim lastRange As Range
    Dim lastColumnName As Variant
    Dim sortRangeAddress As String
    Dim keyIndex As Dictionary
    Dim keyColumnName As String
    Dim sortKeyRangeAddress As String
    
    Set lastRange = targetWorksheet.Cells(lastRow, lastColumn)
    lastColumnName = Split(lastRange.Address(columnAbsolute:=False), "$")(0)
    sortRangeAddress = "A1:" & lastColumnName & lastRow
    Set keyIndex = CreateAssociativeArrayKeyIndex()
    keyColumnName = Split(targetWorksheet.Cells(1, keyIndex("seqNo") + 1).Address(columnAbsolute:=False), "$")(0)
    sortKeyRangeAddress = keyColumnName & "2:" & keyColumnName & lastRow
    With targetWorksheet.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range(sortKeyRangeAddress), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .SetRange Range(sortRangeAddress)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    SortOutputValues = sortRangeAddress
End Function

Private Function GetArrayDimensions(arr As Variant) As Integer()
    Dim result(1) As Integer
    
    On Error Resume Next
    result(0) = UBound(arr, 1) - LBound(arr, 1) + 1
    result(1) = UBound(arr, 2) - LBound(arr, 2) + 1
    On Error GoTo 0

    GetArrayDimensions = result
End Function

