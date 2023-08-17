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

Public Sub SaveWorkbook(wb As Workbook, outputPath As String)
On Error GoTo FINL_L
    Application.DisplayAlerts = False
    wb.SaveAs fileName:=outputPath
    wb.Close
FINL_L:
    Application.DisplayAlerts = True
End Sub

Public Sub CloseWorkbookWithoutSaving(wb As Workbook)
On Error GoTo FINL_L
    Application.DisplayAlerts = False
    wb.Close SaveChanges:=False
FINL_L:
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

Public Function GetSingleExcelFile(folderName As String) As String
On Error GoTo FINL_L
    Dim excelApp As Object
    Dim wb As Workbook
    Dim fileName As String
    Dim fileCount As Integer
    Dim filePath As String
    Dim parentFolder As String
    Dim path As String
    parentFolder = GetParentPath()
    path = GetTargetPath(parentFolder, folderName)
    
    ' Create Excel Application object
    Set excelApp = CreateObject("Excel.Application")
    excelApp.DisplayAlerts = False
    
    ' Check if the specified path exists
    If Dir(path, vbDirectory) = "" Then
        MsgBox "Specified path does not exist."
        GetSingleExcelFile = Empty
    End If
    
    ' Loop through files in the folder
    fileName = Dir(path & "\*.xlsx")
    Do While fileName <> ""
        fileCount = fileCount + 1
        filePath = path & "\" & fileName
        Set wb = excelApp.Workbooks.Open(filePath)
        
        ' Close workbook without saving changes
        wb.Close False
        
        ' Get next file
        fileName = Dir
    Loop
FINL_L:
    ' Clean up
    excelApp.DisplayAlerts = True
    excelApp.Quit
    Set excelApp = Nothing
    
    ' Check file count and return appropriate result
    If fileCount = 0 Then
        MsgBox "No Excel files found in the specified folder."
        GetSingleExcelFile = Empty
    ElseIf fileCount > 1 Then
        MsgBox "Multiple Excel files found in the specified folder."
        GetSingleExcelFile = Empty
    Else
        GetSingleExcelFile = filePath
    End If
End Function

Public Sub GetOrCreateFolder(folderPath As String)
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder (folderPath)
    End If
End Sub
