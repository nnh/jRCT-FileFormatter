Attribute VB_Name = "Consts"
Option Explicit
'
' This module defines global constants.
'
Public Const debugFlag As Boolean = False
Public Const EOF_TEXT As String = "※項目が足らない場合は、適宜行を追加すること。"
Public Const DEFAULT_SHEET_NAME As String = "Sheet1"
Public Const COLUMNNAME_ROW As Integer = 10
Public Const FILLER As String = ""
Public INPUT_LAST_ROW As Integer
Public INPUT_LAST_COLUMN As Integer
Public TARGET_LAST_ROW As Integer
Public TARGET_LAST_COLUMN As Integer
Public OUTPUT_FILENAME As String
Public Function GetDeleteRowValues() As String()
    Const fullWidthString As String = "削除"
    Const halfWidthString As String = "X"
    Const otherString As String = "その他"
    Const postalCodeString As String = "000-0000"
    Const phoneNumberString1 As String = "00-0000-0000"
    Const phoneNumberString2 As String = "000-000-0000"
    Const mailAddressString As String = "X@X.com"
    Const noString As String = "無"
    Const FILLER As String = ""
    Dim res() As String
    ReDim res(TARGET_LAST_COLUMN)
    Dim i As Integer
    i = -1
    ' 0-9
    res(i) = setArrayDeleteRowValues(FILLER, i)
    res(i) = setArrayDeleteRowValues(FILLER, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    res(i) = setArrayDeleteRowValues(halfWidthString, i)
    res(i) = setArrayDeleteRowValues(FILLER, i)
    res(i) = setArrayDeleteRowValues(halfWidthString, i)
    res(i) = setArrayDeleteRowValues(FILLER, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    res(i) = setArrayDeleteRowValues(halfWidthString, i)
    ' 10-19
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    res(i) = setArrayDeleteRowValues(postalCodeString, i)
    res(i) = setArrayDeleteRowValues(otherString, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    res(i) = setArrayDeleteRowValues(phoneNumberString1, i)
    res(i) = setArrayDeleteRowValues(mailAddressString, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    ' 20-30
    res(i) = setArrayDeleteRowValues(postalCodeString, i)
    res(i) = setArrayDeleteRowValues(otherString, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    res(i) = setArrayDeleteRowValues(phoneNumberString2, i)
    res(i) = setArrayDeleteRowValues(phoneNumberString2, i)
    res(i) = setArrayDeleteRowValues(mailAddressString, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    res(i) = setArrayDeleteRowValues(noString, i)
    res(i) = setArrayDeleteRowValues(FILLER, i)
    res(i) = setArrayDeleteRowValues(fullWidthString, i)
    GetDeleteRowValues = res
End Function
Private Function setArrayDeleteRowValues(setText As String, ByRef i As Integer) As String
    i = i + 1
    setArrayDeleteRowValues = setText
End Function
Public Function GetLastRowAndColumnNumber(ws As Worksheet) As Integer()
    Const columnMax As Integer = 100
    Const rowMax As Integer = 10000
    Dim targetLastRange As Range
    Set targetLastRange = ws.Cells(rowMax, columnMax)
    Dim lastRow As Integer
    Dim lastColumn As Integer
    lastRow = 0
    lastColumn = 0
    Dim i As Integer
    Dim j As Integer
    Dim endFlg As Boolean
    endFlg = False
    
    For i = 1 To rowMax
        For j = 1 To columnMax
            If Trim(ws.Cells(i, j).Value) <> "" Then
                If i > lastRow Then
                    lastRow = i
                End If
                If j > lastColumn Then
                    lastColumn = j
                End If
            End If
            If Trim(ws.Cells(i + 1, j).Value) = EOF_TEXT Then
                endFlg = True
                Exit For
            End If
        Next j
        If endFlg Then
            Exit For
        End If
    Next i
    Dim rowAndColumnNumber(2) As Integer
    rowAndColumnNumber(0) = lastRow
    rowAndColumnNumber(1) = lastColumn
    GetLastRowAndColumnNumber = rowAndColumnNumber

End Function
