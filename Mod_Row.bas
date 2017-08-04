Attribute VB_Name = "Mod_Row"
Option Explicit
Option Base 1
Function Row_GetLast(wksCurrentSheet As Worksheet, longColumnNum As Long, Optional boolAddress As Boolean = False) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function returns the last row of data in a column
' first you find the last possible row by version then searches
' backward to the first row that has data in it
' the function returns a row (default) of the address of the last cell
'
' Requirements:
' Cells_ReturnNumberOrLetters()
'
' Inputs:
' wksCurrentSheet
' Type: worksheet
' Desc: worksheet to search
'
' longColumnNum
' Type: long
' Desc: column to search
'
' boolAddress
' Type: boolean
' Desc: flag, row or address
'
' Return:
' Type: long
' Desc: row number
' Type: string
' Desc: cell address
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declarations
Dim rangeCell As Range
Dim longLastRow As Long, longReturnRow As Long

' initialize
longLastRow = wksCurrentSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
longReturnRow = 1

' start
' has something in it by column to the left
Set rangeCell = wksCurrentSheet.Cells(longLastRow, longColumnNum)

' test if cell is empty then go up to find the next
' row cell in column that is not empty
If rangeCell.Value = Empty Then
    Set rangeCell = rangeCell.End(xlUp)
Else
    longReturnRow = rangeCell.Row
End If

' returns the row number or address of cell
If boolAddress = True Then
    ' return the address
    Row_GetLast = Cells_ReturnNumberOrLetters(rangeCell, 3)
Else
    ' return the row number
    Row_GetLast = Cells_ReturnNumberOrLetters(rangeCell, 4)
End If

' reset objects
Set rangeCell = Nothing
End Function
Function Row_Find(stringItemToFind As String, longColumnToSearch As Long, longStartRow As Long, longStopRow As Long) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function will find a row based on the criterea passed in the <stringItemToFind
' field in the column <longColumnToSearch> from row <longStartRow> to row <longStopRow>
' and either return the row the item was found or if not found return zero
'
' Inputs
' stringItemToFind
' Type: string
' Desc: item to find
'
' longColumnToSearch
' Type: long
' Desc: column to search
'
' longStartRow
' Type: long
' Desc: start row
'
' longStopRow
' Type: long
' Desc: row to stop search
'
' Locals
' booleanFoundItem
' Type: boolean
' Desc: flag to determine what to return
'
' Return
' Type: long
' Desc: the row the item is found; if does not find item returns 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' declare
Dim booleanFoundItem As Boolean

' loop variables
Dim a As Long

' initialize
booleanFoundItem = False
If longStartRow = 0 Then longStartRow = 1 Else ' do nothing
If longStopRow = 0 Then longStopRow = 1 Else ' do nothing

' start
For a = longStartRow To longStopRow
    If StrComp(Trim(CStr(Cells(a, longColumnToSearch).Value)), stringItemToFind, vbTextCompare) = 0 Then
        booleanFoundItem = True
        Exit For
    Else
    End If
Next a

' return value
If booleanFoundItem = True Then
    Row_Find = a
Else
    Row_Find = 0
End If
End Function
Sub Row_Copy(wksSource As Worksheet, wksDestination As Worksheet, longCopyRow As Long, longInsertAboveRow As Long)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copies the row from the source worksheet to the destination worksheet and
' turns on the Auto-filter
'
' Inputs
' wksSource
' Type: worksheet
' Desc: source worksheet
'
' wksDestination
' Type: worksheet
' Desc: destination worksheet
'
' longCopyRow
' Type: long
' Desc: row to be copied from source worksheet
'
' longInsertAboveRow
' Type: long
' Desc: row to be for source to be inserted above
'
' Return
' Type:
' Desc:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' object
Dim range_row_source As Range, range_row_dest As Range

' initialize
Set range_row_source = wksSource.Rows(longCopyRow)
Set range_row_dest = wksDestination.Rows(longInsertAboveRow)

' start
range_row_source.Copy
range_row_dest.Insert Shift:=xlDown

' reset objects
Set range_row_source = Nothing
Set range_row_dest = Nothing
End Sub
Function Row_CountVisible(longStartRow As Long, longLastRow As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function takes the last row of a column and the start row
' then counts the cells of the column minus the hidden cells
'
' Inputs:
' longStartRow:
' Type: long
' Desc: the row to start counting
'
' longLastRow:
' Type: long
' Desc: the row to stop counting
'
'
' Return:
' Type: Long
' Desc: the count of cells that are not hidden in the range
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' declare
Dim longCount As Long

' loop variables
Dim a As Long

' initialize
longCount = 0
a = 1

' start
For a = longStartRow To longLastRow
    If Rows(a).Hidden <> True Then
        longCount = longCount + 1
    Else ' do nothing
    End If
Next a

Row_CountVisible = longCount
End Function
Function Row_FindByFirstWord(stringFirstWord As String, longColumnToSearch As Long, longStartRow As Long, longStopRow As Long) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function will find a row based on the criteria passed in the <stringFirstWord>
' field in the column <longColumnToSearch> from row <longStartRow> to row <longStopRow>
' and either return the row the item was found or if not found return zero
'
' Inputs
' stringFirstWord
' Type: string
' Desc: first word to find
'
' longColumnToSearch
' Type: long
' Desc: column to search
'
' longStartRow
' Type: long
' Desc: start row
'
' longStopRow
' Type: long
' Desc: row to stop search
'
' Locals
' booleanFoundItem
' Type: boolean
' Desc: flag to determine what to return
'
' Return
' Type: long
' Desc: the row the item is found; if does not find item returns 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
' declare
Dim intSpace As Integer
Dim stringSplice As String, stringEntry As String
Dim booleanFoundItem As Boolean
 
' loop variables
Dim a As Long
 
' initialize
intSpace = 0
stringSplice = "tsma"
stringEntry = "tsma"
booleanFoundItem = False
 
' start
For a = longStartRow To longStopRow
    stringEntry = Trim(CStr(Cells(a, longColumnToSearch).Value))
    intSpace = InStr(1, stringEntry, " ", vbBinaryCompare)
   
    ' found space and the length of the entry > 0
    If intSpace <> 0 And Len(stringEntry) > 0 Then
        stringSplice = Left(stringEntry, intSpace - 1)
        If StrComp(stringSplice, stringFirstWord, vbTextCompare) = 0 Then
            booleanFoundItem = True
            Exit For
        Else ' do nothing
        End If
    ' only one word in cell, no space but the length is > 0
    ElseIf intSpace = 0 And Len(stringEntry) > 0 Then
        If StrComp(stringEntry, stringFirstWord, vbTextCompare) = 0 Then
            booleanFoundItem = True
            Exit For
        Else ' do nothing
        End If
    Else ' do nothing
    End If
Next a
 
' return value
If booleanFoundItem = True Then
    Row_FindByFirstWord = a
Else
    Row_FindByFirstWord = 0
End If
End Function
Sub Row_FindAndDelete(stringItemToFind As String, longColumnToSearch As Long, longStartRow As Long, longStopRow As Long)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine will find a row based on the criterea passed in the <stringItemToFind
' field in the column <longColumnToSearch> from row <longStartRow> to row <longStopRow>
' and delete that row
'
' Inputs
' stringItemToFind
' Type: string
' Desc: item to find
'
' longColumnToSearch
' Type: long
' Desc: column to search
'
' longStartRow
' Type: long
' Desc: start row
'
' longStopRow
' Type: long
' Desc: row to stop search
'
' Locals
' None
' Type:
' Desc:
'
' Return
' Type: None
' Desc: None
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' declare

' loop variables
Dim a As Long

' initialize
a = longStartRow

' start
Do Until a > longStopRow
    If StrComp(Trim(CStr(Cells(a, longColumnToSearch).Value)), stringItemToFind, vbTextCompare) = 0 Then
        Rows(a).Delete
        Exit Do
    Else
    End If

    ' increment counter
    a = a + 1
Loop
End Sub
Function Row_FindTransition(ByVal longRowStart As Long, ByVal longColumnToSearch As Long, ByVal boolUp As Boolean) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function find the next transition between the next open or cell that has something in it
'
' Requirements:
' None
'
' Inputs:
' longRowStart
' Type: long
' Desc: row of cell
'
' longColumnToSearch
' Type: long
' Desc: column of cell
'
' boolUp
' Type: boolean
' Desc: flag for up or down, if TRUE will go up, if FALSE will go down
'
' Important Info:
' the cell of the transition will be selected
'
' Return:
' variable
' Type: long
' Desc: the row of the transition
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' declare variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' objects
Dim range_temp As Range

' loop
Dim a As Long, b As Long, c As Long
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' initialize variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' objects
Set range_temp = Cells(longRowStart, longColumnToSearch)

' loop
a = 1
b = 1
c = 1
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' begin
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

If boolUp = True Then
    range_temp = range_temp.End(xlUp)
Else
    range_temp = range_temp.End(xlDown)
End If
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' end
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Row_FindTransition = range_temp.Row
End Function
Function Row_GetCriteriaByRow(wks_current As Worksheet, longColumnNumber As Long) As Collection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function will search the row and identify the criteria that is in the column
' and take out the duplicates.  This sub will not count the blanks. Stops at the first
' blank cell.  If the cell is hidden it will not use that cell value in the criteria.
'
' Requirements
' Column_GetCriterea()
' Column_GetLast()
'
' Inputs
' wks_current
' Type: string
' Desc: worksheet's name to work in
'
' longColumnNumber
' Type: long
' Desc: row to be searched
'
' Return
' coll_return
' Type: collection
' Desc: the unique values of each row in the column from the first row to the
'       last row that has data in it
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' declare
Dim coll_return As Collection

' loop variables
Dim a As Long

' initialize
Set coll_return = New Collection

' loop variables
a = 1

' start
For a = 1 To Row_GetLast(wks_current, longColumnNumber)
    coll_return.Add Item:=Row_GetCriteria(wks_current, a, longColumnNumber)
Next a

' return value
Set Row_GetCriteriaByRow = coll_return

' reset objects
Set coll_return = Nothing
End Function

Sub Row_Insert(wksSheet As Worksheet, longNumRowsToInsert As Long, longRow As Long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine will insert a specific number of rows <longNumRowsToInsert> in a worksheet <wksSheet> above a designated row <longRow>
'
' Requirements:
' None
'
' Inputs:
' wksSheet
' Type: worksheet
' Desc: worksheet to insert rows
'
' Important Info:
' None
'
' Return:
' None
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare

' loop variables
Dim a As Long

' initialize

' loop variables
a = 1

' start
For a = 1 To longNumRowsToInsert
    wksSheet.Rows(longRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
Next a
End Sub
Sub Row_Unhide(ByVal longStartRow As Long, ByVal longStopRow As Long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine will unhide all the rows between the start row and stop row
'
' Requirements:
' None
'
' Inputs:
' longStartRow
' Type: long
' Desc: the row to begin unhiding rows
'
' longStopRow
' Type: long
' Desc: the row to stop unhiding rows
'
' Important Info:
' Worksheet must be activated
'
' Return:
' None
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare


' loop variables
Dim a As Long

' initialize

' loop variables
a = 1

' start
For a = longStartRow To longStopRow
    If Rows(a).Hidden = True Then Rows(a).Hidden = False Else ' do nothing
Next a
End Sub
Function Row_GetCriteria(ByVal wksWorksheet As Worksheet, ByVal longRowNumber As Long, Optional ByVal longStartColumn As Long = 1) As Collection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' This Function will search the column and identify the criteria that is in the column
' and take out the duplicates.  This sub will not count the blanks. Stops at the last
' cell of the column.  If the cell is hidden it will not use that cell value in the criterea.
'
'
' requirements:
' Column_GetLast()
'
' Inputs
' collColumnCriterea:
' Type: collection
' Desc: collection to add column values to
'
' longRowNumber:
' Type: long
' Desc: row to be searched
'
' longStartColumn:
' Type: long
' Desc: row to start search
'
' Important Info:
' None
'
' Return
' collReturn
' Type: collection
' Desc: the unique values of the row
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' declare variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Dim collReturn As Collection
Dim longLastColumn As Long

' loop variables
Dim a As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' set objects
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Set collReturn = New Collection

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' initialize variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

longLastColumn = Column_GetLast(wksWorksheet, longRowNumber)

' loop
a = 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' begin
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
 
' turn off errors
On Error Resume Next

' begin search of the column
For a = longStartColumn To longLastColumn
    ' don't use values in hidden rows
    If wksWorksheet.Columns(a).Hidden = False Then
       collReturn.Add Item:=Cells(longRowNumber, a).Value, Key:=Trim(CStr(Cells(longRowNumber, a).Value))
    Else ' do nothing, don't use values in hidden rows
    End If
Next a

' turn errors on
On Error GoTo 0

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' end
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Set Row_GetCriteria = collReturn

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' reset objects
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Set collReturn = Nothing

End Function

