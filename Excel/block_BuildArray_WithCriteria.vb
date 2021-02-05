''' Block: Build String Array with Criteria Start '''
' Description: Use this block of code to build an array from the values in the column of a structured table based on some critera.
' Commit date: 2021-02-05

Dim GenArray() As String '<- adjust the name of the array as necessary
Dim NameOfTable As String
Dim NameOfColumnToCheck As String
Dim NameOfColumnToArray As String
Dim Condition As String
Dim n As String

NameOfTable = "Folders" '<- input name of table
NameOfColumnToCheck = "Include Transmittal Letter" '<- input the name of table column whose values are being checked.
NameOfColumnToArray = "Full Path" '<- input name of table column whose values are to be added to the array.

ColumnReferenceCheck = NameOfTable & "[" & NameOfColumnToCheck & "]"
ColumnReferenceArray = NameOfTable & "[" & NameOfColumnToArray & "]"

' This variable is used to indicate how many times the criteria is met so that the array is redimensioned correctly.
n = -1

For i = 1 To Range(NameOfTable).Rows.Count
    If Range(ColumnReferenceCheck)(i) = True Then '<- update condition as necessary
        n = n + 1
        ReDim Preserve GenArray(n)
        GenArray(n) = Range(ColumnReferenceArray)(i)
    End If
Next i

''' Block: Build String Array with Criteria End '''