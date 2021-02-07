''' Block: Build String Array Start '''
' Description: Use this block of code to build an array from the values in the column of a structured table.
' Commit date: 2021-02-07

'Declare variables
Dim GenArray() As String '<- adjust the name of the array as necessary
Dim NameOfTable As String
Dim NameOfColumn As String
Dim ColumnReference As String

'Input data
NameOfTable = "Folders" '<- input name of table
NameOfColumn = "Full Path" '<- input name of table column.

ColumnReference = NameOfTable & "[" & NameOfColumn & "]"

For i = 1 To Range(NameOfTable).Rows.Count
    ReDim Preserve GenArray(i)
    GenArray(i) = Range(ColumnReference)(i)
Next i

''' Block: Build String Array End '''
