Attribute VB_Name = "CommonUtils"
Option Explicit

Public Function GenerateKey(inputArray As Variant) As Variant
    Dim keys As Dictionary
    Set keys = CreateAssociativeArrayKeyIndex()
    Dim resultArray() As String
    ReDim resultArray(0 To UBound(inputArray, 1))
    
    Dim i As Long
    For i = 0 To UBound(inputArray, 1)
        resultArray(i) = inputArray(i, keys("postalCode")) & inputArray(i, keys("phoneNumber"))
    Next i
    
    GenerateKey = resultArray
End Function

Private Function IsElementInArray(element As Variant, arr As Variant) As Boolean
    Dim count As Long
    count = Application.WorksheetFunction.CountIf(arr, element)
    IsElementInArray = count > 0
End Function

Public Function CreateAssociativeArrayKeyIndex() As Dictionary
    Dim assocArray As Dictionary
    Set assocArray = CreateObject("Scripting.Dictionary")
    
    assocArray("seqNo") = 1
    assocArray("familyName") = 2
    assocArray("givenName") = 3
    assocArray("facilityName") = 8
    assocArray("postalCode") = 11
    assocArray("facilityAddress") = 13
    assocArray("phoneNumber") = 14
    
    Set CreateAssociativeArrayKeyIndex = assocArray
End Function

