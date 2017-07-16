Attribute VB_Name = "Mod_Column"
Option Explicit
Option Base 1
Function Column_CopyPaste(wks_source As Worksheet, string_column_source As String, long_column_source As Long, _
                        wks_dest As Worksheet, long_column_dest As Long, long_row_dest_start As Long _
                        , bool_header_row As Boolean)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function copies data from a column in a source worksheet to the destintion
' worksheet then returns the last row in that column
'
' Requirements:
' Row_GetLast()
'
' Inputs:
' stringSourceWorksheet
' Type: string
' Desc: source worksheet name
'
' stringCopyColumn
' Type: string
' Desc: column to copy data from the source worksheet, column string destination
'
' longCopyColumn
' Type: long
' Desc: source column on destination worksheet, column number
'
' stringDestinationWorksheet
' Type: string
' Desc: destination worksheet name
'
' longDesinationColumn
' Type: long
' Desc: destination column on destination worksheet
'
' longDestinationStartRow
' Type: long
' Desc: the last row of the column in the destination worksheet that the data was
'       copied to
'
' boolHeaderRow
' Type: boolean
' Desc: flag to tell function if there is a header row in source worksheet
'
' Return:
' Type: Long
' Desc: the last row of the destination column the source data was copied to
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' declare
Dim range_copy As Range, range_paste As Range
Dim long_source_last_row As Long
Dim string_copy As String

' initialize
long_source_last_row = Row_GetLast(wks_source, long_column_source)
string_copy = "tsma"

' start
' check for header and gen copy string
If bool_header_row = True Then
    string_copy = string_column_source & "2:" & string_column_source & CStr(long_source_last_row)
Else
    string_copy = string_column_source & "1:" & string_column_source & CStr(long_source_last_row)
End If

' copy column
Set range_copy = wks_source.Range(string_copy)
range_copy.Copy

' past to destination
Set range_paste = wks_dest.Cells(long_row_dest_start, long_column_dest)
range_paste.PasteSpecial Paste:=xlPasteValues, SkipBlanks:=True
Application.CutCopyMode = False

' return value, last row of destination
Column_CopyPaste = Row_GetLast(wks_dest, long_column_dest)
End Function
Function Column_GetLast(ByVal wksCurrentSheet As Worksheet, ByVal longRowNum As Long, Optional ByVal intReturnType As Integer = 1) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function returns the last column of data in a row
' first you find the last possible column by version then searches
' backward to the first column that has data in it
' the function returns a column in a designated type
'
' Requirements:
' Cells_ReturnNumberOrLetters()
'
' Inputs:
' wksCurrentSheet
' Type: worksheet
' Desc: worksheet to search
'
' longRowNum
' Type: long
' Desc: row to search
'
' intReturnType
' Type: integer
' Desc: flag, 1 = column (long), 2 = column letter (string), 3 = address of cell (string)
'
' Return:
' Type: long
' Desc: row number
' Type: string
' Desc: cell letter
' Type: string
' Desc: cell address
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declarations
Dim rangeCell As Range
Dim longLastCol As Long

' initialize
longLastCol = wksCurrentSheet.Cells.SpecialCells(xlCellTypeLastCell).Column

' start
' sets range object to the last cell in the row
Set rangeCell = wksCurrentSheet.Cells(longRowNum, longLastCol)

' tests if cell is empty then finds the
' next cell that isn't empty
If rangeCell.Value = Empty Then
    Set rangeCell = rangeCell.End(xlToLeft)
Else ' do nothing
End If

' return
Column_GetLast = Cells_ReturnNumberOrLetters(rangeCell, intReturnType)

' reset objects
Set rangeCell = Nothing
End Function

Function Column_FindByTimePeriod(ByVal wksWorksheet As Worksheet, ByVal longRowToSearch As Long, ByVal intTimePeriod As Integer, _
                                Optional intReturnType As Integer = 1) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function will search a row of dates from the last column with an entry to find the column that equals the time period passed
' (in days) by subtracting the last column date with the searched column date;
'
' Requirements:
' Column_GetLast()
' Cells_ReturnNumberOrLetters()
'
' Inputs:
' wksWorksheet
' Type: worksheet
' Desc: the worksheet to search
'
' longRowToSearch
' Type: long
' Desc: the row to search for the column based on the time period
'
' intTimePeriod
' Type: integer
' Desc: the time period to look for the correct column (in days)
'
' intReturnType
' Type: integer
' Desc: the flag to choose a return value, number or letter;
'       the default is the column number
'
' Important Info:
' The row is formatted as the oldest entry is in column 1 and the newest entry is in the last column
'
' Return:
' variable
' Type: variant
' Desc: the start column based on the date the time period passed in days;
'       if there the column is not found then the first column will be returned
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' declare variables
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Dim longColTest As Long, longColReturn As Long
Dim intPosit As Integer
Dim stringCellAddress As String
Dim variantReturnValue As Variant
Dim dateTest As Date, dateNewest As Date
Const boolError As Boolean = False

' loop
Dim a As Long, b As Long, c As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' set objects
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' initialize variables
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

longColTest = Column_GetLast(wksWorksheet, longRowToSearch)
longColReturn = 1
intPosit = 0
stringCellAddress = "tsma"
dateTest = #1/11/1111#
If IsDate(Cells(longRowToSearch, longColTest).Value) = True Then
    dateNewest = Cells(longRowToSearch, longColTest).Value
Else ' do nothing
End If

' loop
a = 1
b = 1
c = 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' begin
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' search the row from last column to the first column
Do Until longColTest < 1
    ' test to ensure the value is a date
    If IsDate(Cells(longRowToSearch, longColTest).Value) = True Then
        ' get date
        dateTest = Cells(longRowToSearch, longColTest).Value

        ' test date difference
        If DateDiff("d", dateTest, dateNewest) >= intTimePeriod Then
            ' return the previous column
            longColReturn = longColTest

            ' exit loop
            Exit Do
        Else ' do nothing
        End If
    Else ' do nothing
    End If

    ' reset variable
    dateTest = Empty

    ' decrement column
    longColTest = longColTest - 1

Loop

' get the column number or letter
variantReturnValue = Cells_ReturnNumberOrLetters(Cells(longRowToSearch, longColReturn), intReturnType)

' error handling example
'On error goto <error ID Label>:
'If …. Then
'
'Else
'<resume code label>:
' on error goto 0
'End if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' error handling
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

If boolError = True Then
'<error ID label>:
'    error handling code
'    Resume <resume code label>: ' goto <resume code label> to resume code
Else ' do nothing
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' end
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' return value
Column_FindByTimePeriod = variantReturnValue

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' reset objects
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
End Function

Function Column_Find(stringItemToFind As String, longRowToSearch As Long, longStartCol As Long, longStopCol As Long, Optional intReturnType As Integer = 1) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function will find a column based on the criteria passed in the <stringItemToFind>
' field in the row <longRowToSearch> from row <longStartCol> to row <longStopcol>
' and either return the column (letter or number) the item was found or if not found return zero
'
' Inputs
' stringItemToFind
' Type: string
' Desc: item to find
'
' longRowToSearch
' Type: long
' Desc: roe to search
'
' longStartCol
' Type: long
' Desc: start column
'
' longStopCol
' Type: long
' Desc: column to stop search
'
' intReturnType
' Type: integer
' Desc: the flag to determine the return type
'
' Locals:
' booleanFoundItem
' Type: boolean
' Desc: flag to determine what to return
'
' Return
' variable
' Type: variant
' Desc: the column # or letter if the item is found; if does not find item returns 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' declare
Dim booleanFoundItem As Boolean
Dim stringCol As String
Dim intPosit As Integer
Dim variantReturnValue As Variant

' loop variables
Dim a As Long

' initialize
booleanFoundItem = False
stringCol = "tsma"
intPosit = 0
variantReturnValue = 0

' start
For a = longStartCol To longStopCol
    If StrComp(Trim(CStr(Cells(longRowToSearch, a).Value)), stringItemToFind, vbTextCompare) = 0 Then
        booleanFoundItem = True
        Exit For
    Else ' do nothing
    End If
Next a

' get return value
If booleanFoundItem = True Then
    Select Case intReturnType
        Case 1: ' column #
            variantReturnValue = a
        Case 2: ' column letter
            stringCol = Cells(longRowToSearch, a).Address(True, False)
            intPosit = InStr(1, stringCol, "$")
            variantReturnValue = Left(stringCol, intPosit - 1)
        Case Else ' do nothing
    End Select
Else ' do nothing
End If

' return value
Column_Find = variantReturnValue
End Function
Function Column_FindByDate(dateReference As Date, longRow As Long, longStartCol As Long, longStopCol As Long, intReturnType As Integer, boolDay As Boolean, _
                          boolMonth As Boolean, boolYear As Boolean) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function finds the column by date then returns that column
'
' Requirements:
' None
'
' Inputs:
' dateReference
' Type: date
' Desc: the date that needs to be found
'
' longRow
' Type: long
' Desc: the row to search
'
' longStartCol
' Type: long
' Desc: column to start search
'
' longStartCol
' Type: long
' Desc: column to stop search
'
' intReturnType
' Type: integer
' Desc: a number of letter(s) returned (1 or 2)
'
' boolDay
' Type: boolean
' Desc: flag to use day
'
' boolMonth
' Type: boolean
' Desc: flag to use month
'
' boolYear
' Type: date
' Desc: flag to use year
'
' Important Info:
' worksheet to search must be activated or selected
'
' Return:
' variable
' Type: variant
' Desc: the column letter or column
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' declare variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
Dim variantReturnColumn As Variant
Dim intDayRef As Integer, intMonthRef As Integer, intYearRef As Integer, intDayTest As Integer, intMonthTest As Integer, intYearTest As Integer
Dim intPosit As Integer
Dim stringCol As String
Dim dateTest As Date
 
' loop
Dim a As Long, b As Long, c As Long
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' initialize variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
variantReturnColumn = 0
intDayRef = 0
intMonthRef = 0
intYearRef = 0
intDayTest = 0
intMonthTest = 0
intYearTest = 0
dateTest = #1/11/1111#
 
' loop
a = 1
b = 1
c = 1
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' begin
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
 
' get integers of date parts
If boolDay = True Then intDayRef = Day(dateReference) Else ' do nothing
If boolMonth = True Then intMonthRef = Month(dateReference) Else ' do nothing
If boolYear = True Then intYearRef = Year(dateReference) Else ' do nothing
 
' loop through columns in row
For a = longStartCol To longStopCol
    If IsDate(Cells(longRow, a).Value) = True Then
        ' get test date
        dateTest = Cells(longRow, a).Value
       
        ' get test integers
        If boolDay = True Then intDayTest = Day(dateTest) Else ' do nothing
        If boolMonth = True Then intMonthTest = Month(dateTest) Else ' do nothing
        If boolYear = True Then intYearTest = Year(dateTest) Else ' do nothing
       
        ' test dates
        If boolDay = True And boolMonth = True And boolYear = True Then
            If intDayTest = intDayRef And intMonthTest = intMonthRef And intYearTest = intYearRef Then variantReturnColumn = a Else ' do nothing
        ElseIf boolDay = True And boolMonth = True And boolYear = False Then
            If intDayTest = intDayRef And intMonthTest = intMonthRef Then variantReturnColumn = a Else ' do nothing
        ElseIf boolDay = True And boolMonth = False And boolYear = True Then
            If intDayTest = intDayRef And intYearTest = intYearRef Then variantReturnColumn = a Else ' do nothing
        ElseIf boolDay = False And boolMonth = True And boolYear = True Then
            If intMonthTest = intMonthRef And intYearTest = intYearRef Then variantReturnColumn = a Else ' do nothing
        ElseIf boolDay = True And boolMonth = False And boolYear = False Then
            If intDayTest = intDayRef Then variantReturnColumn = a Else ' do nothing
        ElseIf boolDay = False And boolMonth = True And boolYear = False Then
            If intMonthTest = intMonthRef Then variantReturnColumn = a Else ' do nothing
        ElseIf boolDay = False And boolMonth = False And boolYear = True Then
            If intYearTest = intYearRef Then variantReturnColumn = a Else ' do nothing
        Else ' do nothing
        End If
    Else ' do nothing
    End If
Next a
 
' get return value
Select Case intReturnType
    Case 1: ' column #
        variantReturnColumn = Cells(longRow, variantReturnColumn).Column
    Case 2: ' column letter
        stringCol = Selection.Address(False, False)
        intPosit = InStr(1, stringCol, CStr(longRow))
        If Len(stringCol) = intPosit Then
            variantReturnColumn = Left(stringCol, Len(stringCol) - intPosit + 1)
        Else
            variantReturnColumn = Left(stringCol, Len(stringCol) - intPosit)
        End If
    Case Else ' do nothing
End Select

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' end
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
Column_FindByDate = variantReturnColumn
 
End Function
Function Column_GetCriteria(ByVal wksWorksheet As Worksheet, ByVal longColumnNumber As Long, Optional ByVal longStartRow As Long = 1) As Collection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' This Function will search the column and identify the criteria that is in the column
' and take out the duplicates.  This sub will not count the blanks. Stops at the last
' cell of the column.  If the cell is hidden it will not use that cell value in the criterea.
'
'
' requirements:
' Row_GetLast()
'
' Inputs
' collColumnCriterea:
' Type: collection
' Desc: collection to add column values to
'
' longColumnNumber:
' Type: long
' Desc: column to be searched
'
' longStartRow:
' Type: long
' Desc: row to start search
'
' Important Info:
' None
'
' Return
' collReturn
' Type: collection
' Desc: the unique column values
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' declare variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Dim collReturn As Collection
Dim longLastRow As Long

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

longLastRow = Row_GetLast(wksWorksheet, longColumnNumber)

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
For a = longStartRow To longLastRow
    ' don't use values in hidden rows
    If Rows(a).Hidden = False Then
       collReturn.Add Item:=Cells(a, longColumnNumber).Value, Key:=Trim(CStr(Cells(a, longColumnNumber).Value))
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

Set Column_GetCriteria = collReturn

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' reset objects
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Set collReturn = Nothing

End Function
Function Column_GetCriteriaByColumn(wks_current As Worksheet, longRowNumber As Long) As Collection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function will search the column and identify the criteria that is in the column
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
' Desc: contains collections; the unique values of each column from the first column
'       to the last column with data in it
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' declare
Dim coll_return As Collection
Dim booleanFirstValue As Boolean, booleanSucessfullAdd As Boolean

' loop variables
Dim a As Long

' initialize
Set coll_return = New Collection

' loop variables
a = 1

' start
For a = 1 To Column_GetLast(wks_current, longRowNumber)
    coll_return.Add Item:=Column_GetCriteria(wks_current, a, longRowNumber)
Next a

' return value
Set Column_GetCriteriaByColumn = coll_return

' reset objects
Set coll_return = Nothing
End Function

Function Column_InsertWithHeader(stringHeader As String, longColumn As Long, stringNumberFormat) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function inserts a column with a header name to the left of
' the desired column, formats that column to a number with no decimal
' places and retuns the next column
'
' Requirements:
' none
'
' Inputs:
' stringHeader
' Type: string
' Desc: the column name
'
' longColumn
' Type: long
' Desc: the column to insert the new column before
'
' stringNumberFormat
' Type: string
' Desc: the number fomat for the column
'
' Return:
' Type: long
' Desc: the next column
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare
Dim col_temp As Range

' initialize

' start
Set col_temp = Columns(longColumn)
col_temp.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
col_temp.NumberFormat = stringNumberFormat

' header cell
With Cells(1, longColumn)
    .NumberFormat = "@"
    .Value = stringHeader
End With

' return the next column
Column_InsertWithHeader = longColumn + 1
End Function
Function Column_TakeOutBlankCells(wks_current As Worksheet, longColumn As Long, longStartRow As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function will look for empty cells in an indicated column and delete that row
' it will also return the new last row
'
' Requirements:
' Row_GetLast()
'
' Inputs:
' wks_current
' Type: worksheet
' Desc: orksheet of column to look in
'
' longColumn
' Type: long
' Desc: column number to search for empty rows
'
' longStartRow
' Type: long
' Desc: row to begin search
'
' Important Info:
'
' Return:
' Type: long
' Desc: new last row of column after rows are deleted
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare
Dim longLastRow As Long

' loop variables
Dim a As Long, b As Long, c As Long

' initialize
longLastRow = Row_GetLast(wks_current, longColumn)
a = longStartRow

' start
Do Until a > longLastRow
    If wks_current.Cells(a, longColumn).Value2 = Empty Then wks_current.Rows(a).Delete Else ' do nothing

    ' increment counter
    a = a + 1
Loop

' return value
Column_TakeOutBlankCells = Row_GetLast(wks_current, longColumn)
End Function

