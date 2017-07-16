Attribute VB_Name = "Mod_Search"
Option Explicit
Option Base 1
Function Search_Log(wksSheet As Worksheet, longColumn As Long, variantValue As Variant, booleanString As Boolean) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function searches a worksheets column by a simple logrithmic algorithm
' for a specific value.
'
' ** Note: must have a header row
'
' Requirements:
' Row_GetLast()
' AutoFilter_Clear()
'
' wksSheet:
' Type: worksheet
' Desc: worksheet data to search is in
'
' longColumn
' Type: long
' Desc: column to be searched
'
' variantValue
' Type: variant
' Desc: value searching for
'
' booleanString
' Type: boolean
' Desc: flag if <varaintValue> is string
'
' Return:
' variable
' Type: variant
' Desc: return false if not detected, worksheet row if found
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' declare
Dim longLowerRow As Long, longUpperRow As Long, longTestRow As Long, longMaxLoop As Long, longCounter As Long
Dim int_dollar_posit As Integer
Dim stringColumn As String
Dim variantReturnValue As Variant

' loop variables
Dim a As Long

' initialize
longLowerRow = 2
longUpperRow = 2
longTestRow = 2
longMaxLoop = 1
longCounter = 0
stringColumn = "tsma"
variantReturnValue = False

' loop variable
a = 1

' start
' get the column string
stringColumn = Left(Cells(1, longColumn).Address(RowAbsolute:=True, ColumnAbsolute:=False), 1)
int_dollar_posit = InStr(1, stringColumn, "$")
stringColumn = Left(stringColumn, int_dollar_posit - 1)

' short worksheet by column
Call AutoFilter_Clear
wksSheet.Autofilter.Sort.SortFields.Add Key:=Range(stringColumn & ":" & stringColumn), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With wksSheet.Autofilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

' get last row and max number of loops
longUpperRow = Row_GetLast(wksSheet, longColumn)
longMaxLoop = Int((Log(CDbl(longUpperRow - longLowerRow + 1)) / Log(2#)) + 1)

' start loop
For a = 1 To longMaxLoop
    ' get the row to test
    longTestRow = Int((longUpperRow - longLowerRow) / 2) + longLowerRow
    
    ' special case; only two rows to test
    If longUpperRow - longLowerRow = 1 Then
        If booleanString = True Then
            If StrComp(Trim(Cells(longTestRow, longColumn).Value), variantValue, vbTextCompare) = 0 Then
                variantReturnValue = longTestRow
            ElseIf StrComp(Trim(Cells(longTestRow + 1, longColumn).Value), variantValue, vbTextCompare) = 0 Then
                variantReturnValue = longTestRow
            Else ' do nothing
            End If
            Exit For
        Else
            If Cells(longTestRow, longColumn).Value = variantValue Then
                variantReturnValue = longTestRow
            ElseIf Cells(longTestRow + 1, longColumn).Value = variantValue Then
                longLowerRow = longTestRow
            Else ' do nothing
            End If
            Exit For
        End If
    Else ' conduct test
        ' test value in cell
        If booleanString = True Then
            If StrComp(Trim(Cells(longTestRow, longColumn).Value), variantValue, vbTextCompare) = 0 Then
                variantReturnValue = longTestRow
                Exit For
            ElseIf StrComp(Trim(Cells(longTestRow, longColumn).Value), variantValue, vbTextCompare) = -1 Then
                longUpperRow = longTestRow
            ElseIf StrComp(Trim(Cells(longTestRow, longColumn).Value), variantValue, vbTextCompare) = 1 Then
                longLowerRow = longTestRow
            Else ' do nothing
            End If
        Else
            If Cells(longTestRow, longColumn).Value = variantValue Then
                variantReturnValue = longTestRow
                Exit For
            ElseIf Cells(longTestRow, longColumn).Value < variantValue Then
                longLowerRow = longTestRow
            ElseIf Cells(longTestRow, longColumn).Value > variantValue Then
                longUpperRow = longTestRow
            Else ' do nothing
            End If
        End If
    End If
Next a

' return value
Search_Log = variantReturnValue
End Function

