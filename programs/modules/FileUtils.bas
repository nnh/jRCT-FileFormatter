Attribute VB_Name = "FileUtils"
Option Explicit
'
' This module defines functions related to file operations.
'
Public Function GetTargetPath(parentPath As String, folderName As String) As String
    GetTargetPath = parentPath & "\" & folderName
End Function

Public Function GetParentPath() As String
    Dim fso As Scripting.FileSystemObject
    Dim workbookPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    workbookPath = ThisWorkbook.path
    GetParentPath = fso.GetParentFolderName(workbookPath)
End Function

Public Function OpenWorkbook(folderName As String, filename As String) As Workbook
    Dim parentFolder As String
    Dim inputFolder As String
    parentFolder = GetParentPath()
    inputFolder = GetTargetPath(parentFolder, folderName)
    Set OpenWorkbook = Workbooks.Open(inputFolder & "\" & filename)
End Function

Public Sub SaveWorkbook(wb As Workbook, outputPath As String)
On Error GoTo finl_l
    Application.DisplayAlerts = False
    wb.SaveAs filename:=outputPath
    wb.Close
finl_l:
    Application.DisplayAlerts = True
End Sub

Public Sub CloseWorkbookWithoutSaving(wb As Workbook)
On Error GoTo finl_l
    Application.DisplayAlerts = False
    wb.Close SaveChanges:=False
finl_l:
    Application.DisplayAlerts = True
End Sub

Public Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    Set wb = ThisWorkbook
    sheetExists = False
    
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    If Not sheetExists Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = sheetName
    Else
        Set ws = wb.Worksheets(sheetName)
    End If
    
    Set GetOrCreateSheet = ws
End Function

