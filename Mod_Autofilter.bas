Attribute VB_Name = "Mod_Autofilter"
Option Explicit
Option Base 1
Sub AutoFilter_Clear(Optional stringWksName As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine clears the autofilter of values and checks to see if there is a filter
'
' Requirements:
' None
'
' Inputs:
' stringWksName
' Type: string
' Desc: worksheet name to activate
'
' Important Info:
' sheet must be activated or selected
'
' Return:
' None
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare
Dim wks_current As Worksheet
Dim filtersCollection As Filters
Dim filterLoop As Filter

' start
' activate sheet if passed
If stringWksName = Empty Then
    Set wks_current = Worksheets(ActiveSheet.Name)
Else
    Set wks_current = Worksheets(stringWksName)
End If

' turn on autofilter & filter newest to oldest
If IsEmpty(Range("A1").Value) = False Then
    If wks_current.Range("A1").AutoFilterMode = False Then wks_current.Range("A1").Autofilter Else ' do nothing
   
    ' clear filter values
    ' clear all filter values and show all data
    If wks_current.AutoFilterMode = True Then
        Set filtersCollection = wks_current.Autofilter.Filters
        If filtersCollection.Count > 0 Then
            For Each filterLoop In filtersCollection
                If filterLoop.On = True Then
                    wks_current.ShowAllData
                    Exit For
                Else
                End If
            Next
        Else
        End If
    Else
    End If
Else
End If
End Sub
Sub AutoFilter_OnOff(wksWorksheet As Worksheet, rngCell As Range, boolTurnOn As Boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine will turn the autofilter on or off
'
' Requirements:
' None
'
' Inputs:
' wksWorksheet
' Type: worksheet object
' Desc: the worksheet to turn the autofilter on or off
'
' rngCell
' Type: range object, worksheet cell
' Desc: the cell to toggle the autofilter
'
' boolTurnOn
' Type: boolean
' Desc: flag to turn the autofilter on or off
'
' Important Info:
' None
'
' Return:
' None
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare objects
Dim wks_autofilter As Autofilter
Dim range_cell As Range

' set objects
Set wks_autofilter = wksWorksheet.Autofilter
Set range_cell = wksWorksheet.Range(rngCell.Address(False, False))

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' begin
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' if autofilter is on and it should be on
If Not wks_autofilter Is Nothing And boolTurnOn = True Then
    ' do nothing

' if autofilter is off and it should be on
ElseIf wks_autofilter Is Nothing And boolTurnOn = True Then
    range_cell.Autofilter

' if autofilter is on and it should be off
ElseIf Not wks_autofilter Is Nothing And boolTurnOn = False Then
    range_cell.Autofilter
Else ' do nothing
End If

' reset objects
Set wks_autofilter = Nothing
Set range_cell = Nothing
End Sub
Sub Autofilter_SingleSort(wksWorksheet As Worksheet, stringSortRange As String, boolDescending As Boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine will sort a worksheet from a stringSortRange.  this sub will only sort on one sort field not multiple
'
' Requirements:
' ClearAutoFilter()
'
' Inputs:
' wksWorksheet
' Type: worksheet
' Desc: worksheet to be sorted
'
'
' stringSortRange
' Type: string
' Desc: range to be sorted
'
' boolDescending
' Type: boolDescending
' Desc: True = descending, False = ascending
'
' Important Info:
' this subroutine only sorts one sort field
' this subroutine will clear all the sortfields
'
' Return:
' None
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare
Dim varaintOrder As Variant

' initialize
varaintOrder = xlDescending

' start
' test for ascending
If boolDescending = False Then varaintOrder = xlAscending Else ' do nothing

' clear autofilter
Call AutoFilter_Clear(wksWorksheet.Name)

' begin sort
wksWorksheet.Rows.Hidden = False
wksWorksheet.Autofilter.Sort.SortFields.Clear
wksWorksheet.Autofilter.Sort.SortFields.Add Key:=Range(stringSortRange), SortOn:=xlSortOnValues, Order:=varaintOrder, DataOption:=xlSortNormal

With wksWorksheet.Autofilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
End Sub
