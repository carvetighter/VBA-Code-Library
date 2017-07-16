Attribute VB_Name = "Mod_PowerPoint"
Option Explicit
Option Base 1
Sub PowerPoint_CutPasteChartsFromExcel(variantCharts As Variant, ppSlide As Slide, wksExcel As Worksheet)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine will look for charts in PowerPoint and Excel then cut and past charts as appropriate based on the chart titles in
' the array varinatCharts()
'
' Requirements:
' None
'
' Inputs:
' variantCharts()
' Type: variant
' Desc: string array with the names of the charts (chart titles)
' variantCharts(x,1) -> chart title to be cut from power point slide
' variantCharts(x,2) -> chart title to find in excel copy & past to power point
'
' ppSlide
' Type: Slide
' Desc: the PowerPoint slide object
'
'
' wksExcel
' Type: worksheet
' Desc: the excel worksheet object to find the chart cut & paste to PowerPoint
'
' Important Info:
' None
'
' Return:
' None
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare
Dim doubleChartHeight As Double, doubleChartWidth As Double, doubleChartTop As Double, doubleChartLeft As Double
Dim stringExcelChartTitle As String
Dim shapeTemp As Object, chartTemp As Object

' loop variables
Dim a As Long, b As Long, c As Long, d As Long

' initialize
doubleChartHeight = 0
doubleChartWidth = 0
doubleChartTop = 0
doubleChartLeft = 0
stringExcelChartTitle = "tsma"

' loop variables
a = 1
b = 1
c = 1
d = 1

' start
For c = 1 To UBound(variantCharts, 1)
    For d = 1 To ppSlide.Shapes.Count
        Set shapeTemp = ppSlide.Shapes.Item(d)

        If PowerPoint_IDShape(shapeTemp) = 1 Then
            If StrComp(CStr(variantCharts(c, 1)), shapeTemp.Chart.ChartTitle.Caption, vbTextCompare) = 0 Then
                ' get PowerPoint shape, size and dimentions
                doubleChartHeight = shapeTemp.Height
                doubleChartWidth = shapeTemp.Width
                doubleChartTop = shapeTemp.Top
                doubleChartLeft = shapeTemp.Left
                shapeTemp.Delete
               
                ' search through excel shapes
                For b = 1 To wksExcel.Shapes.Count
                    ' ID's a chart
                    If PowerPoint_IDShape(wksExcel.Shapes.Item(b)) = 1 Then
                       
                        ' ID's the correct chart
                        stringExcelChartTitle = CStr(variantCharts(c, 2))
                        Set chartTemp = wksExcel.Shapes.Item(b).Chart
                        If StrComp(chartTemp.ChartTitle.Caption, stringExcelChartTitle, vbTextCompare) = 0 Then
                            ' copy's and pastes the chart to PowerPoint with the size and dimentions of the previous
                            ' chart
                            wksExcel.Shapes.Item(b).Copy
                            ppSlide.Shapes.Paste
                            Set shapeTemp = ppSlide.Shapes(ppSlide.Shapes.Count)
                            With shapeTemp
                                .Height = doubleChartHeight
                                .Width = doubleChartWidth
                                .Top = doubleChartTop
                                .Left = doubleChartLeft
                            End With
                           
                            ' clear clipboard
                            Application.CutCopyMode = False
                           
                            ' clean up
                            Set shapeTemp = Nothing
                            Set chartTemp = Nothing
                            Exit For ' loop "b"
                        Else ' do  nothing
                        End If
                    Else
                    End If
                Next b
                Exit For ' loop "d"
            Else
            End If
        Else
        End If
    Next d
Next c
End Sub
Function PowerPoint_IDShape(objShape As Object)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function identifies the type of shape between chart, table, textframe and returns and id number
'
' Requirements:
' None
'
' Inputs:
' objShape
' Type: object/shape
' Desc: the object which is a shape
'
' Important Info:
' shape ID's are below
' 1 -> chart
' 2 -> table
' 3 -> text frame
'
' Return:
' Type: integer
' Desc: the integer will indicate the type of shape
' 1 -> chart
' 2 -> table
' 3 -> text frame
' 0 -> shape not chart, table, text frame
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare
Dim integerShapeID As Integer

' initialize
integerShapeID = 0

' start
If objShape.HasChart = msoTrue Then
    integerShapeID = 1
ElseIf objShape.HasTable = msoTrue Then
    integerShapeID = 2
ElseIf objShape.HasTextFrame = msoTrue Then
    integerShapeID = 3
Else ' do nothing
End If

PowerPoint_IDShape = integerShapeID
End Function
Function PowerPoint_ModifyTitle(stringSlideTitle As String, dateReportDate As Date, longSlideNumber As Long) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine add the date to the title of the slide
'
' Requirements:
' None
'
' Inputs:
' stringSlideTitle
' Type: string
' Desc: the string of the title portion of the slide
'
' dateReportDate
' Type: date
' Desc: the date the report covers, end date of the report
'
' longSlideNumber
' Type: long
' Desc: the number of the slide in the PowerPoint presentation
'
' Important Info:
' None
'
' Return:
' <variable>
' Type: string
' Desc: title with new date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare
Dim longStringPosit As Long
Dim stringTemp As String

' initialize
longStringPosit = 0
stringTemp = "tsma"

' start
longStringPosit = InStr(1, stringSlideTitle, Chr(11), vbBinaryCompare)
stringTemp = Mid(stringSlideTitle, 1, longStringPosit)

If longSlideNumber = 3 Then
    stringTemp = stringTemp & CStr(Format(dateReportDate, "dddd, mmmm dd, yyyy")) & " by Cycle Time"
Else
    stringTemp = stringTemp & CStr(Format(dateReportDate, "dddd, mmmm dd, yyyy"))
End If

PowerPoint_ModifyTitle = stringTemp
End Function
Sub PowerPoint_AppendixSlide(ppTempSlide As PowerPoint.Slide, dateReportDate As Date)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine will create slide 3 from the PowerPoint template and the data from the excel spreadsheet.  It is looking for all
' the activated contracts for the week of the newest date (dateReportDate) as an input into this sub.  It will add all those
' contracts to slide 3.
'
' Requirements:
' ModifyTitle() <function>
' IDShape() <function>
' TestGroupItems() <function>
'
'
' Inputs:
' ppTempSlide
' Type: PowerPoint.Slide
' Desc: the object "slide" from PowerPoint library
'
' dateReportDate
' Type: date
' Desc: date the report is through
'
' Important Info:
' None
'
' Return:
' None
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare
Dim boolModTitle As Boolean
Dim collShapeIds As Collection
Dim shapePowerPoint As Object, shapeTemp As Object

' loop variables
Dim a As Long, b As Long, c As Long

' initialize
boolModTitle = False
Set collShapeIds = New Collection

' loop variables
a = 1
b = 1
c = 1

' start
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' add information to table on slide
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' determine if there are group items
' add all shapes in power point slide to collection
For b = 1 To ppTempSlide.Shapes.Count
    Set shapePowerPoint = ppTempSlide.Shapes.Item(b)
    If PowerPoint_TestGroupItems(shapePowerPoint) = True Then
        For c = 1 To shapePowerPoint.GroupItems.Count
            collShapeIds.Add Item:=shapePowerPoint.GroupItems.Item(c)
        Next c
    Else ' do nothing
        collShapeIds.Add Item:=shapePowerPoint
    End If
Next b
Set shapePowerPoint = Nothing

' add information to PowerPoint slide
For a = 1 To collShapeIds.Count
    Set shapeTemp = collShapeIds.Item(a)
   
    Select Case PowerPoint_IDShape(shapeTemp)
        Case 1 ' chart
            ' do nothing, no chart on this slide
        Case 2 ' table
            ' do nothing, no table on slide
        Case 3 ' textframe
            If StrComp(Left(shapeTemp.Name, 5), "Title", vbTextCompare) = 0 And boolModTitle = False Then
                shapeTemp.TextFrame.TextRange.Text = PowerPoint_ModifyTitle(shapeTemp.TextFrame.TextRange.Text, dateReportDate, ppTempSlide.SlideNumber)
                shapeTemp.TextFrame.VerticalAnchor = msoAnchorMiddle
                boolModTitle = True
            Else ' do nothing
            End If
        Case Else
    End Select
   
    ' clean up
    Set shapeTemp = Nothing
Next a
End Sub
Function PowerPoint_TestGroupItems(ByVal shapeTest As Object) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function tests for group items in a PowerPoint slide
'
' Requirements:
' none
'
' Inputs:
' shapeTest
' Type: object/shape
' Desc: this is a shape from a PowerPoint slide
'
' Important Info:
' None
'
' Return:
' variable
' Type: boolean
' Desc: answer to whether there are grouped shapes in the shape object to be tested
' false -> no group shapes
' true -> there are group shapes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' start
On Error GoTo errorHandler
shapeTest.GroupItems
PowerPoint_TestGroupItems = True
On Error GoTo 0
Exit Function

errorHandler:
PowerPoint_TestGroupItems = False
On Error GoTo 0
End Function

