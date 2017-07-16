Attribute VB_Name = "Mod_Charts"
Option Explicit
Option Base 1
Sub chart_new(ByVal wksGraph As Worksheet, ByVal wksData As Worksheet, ByVal longRowDataHeader As Long, ByVal typeGraph As XlChartType, _
              Optional ByVal longSeriesColor As Long = 255, Optional ByVal longColStart As Long = 1, Optional ByVal longColStop As Long = 10, _
              Optional ByVal longGraphTopPosit As Long = 1, Optional ByVal longGraphLeftPosit As Long = 1, Optional ByVal longGraphWidth As Long = 50, _
              Optional ByVal longGraphHeight As Long = 50, Optional ByVal longMarkerColor As Long = 16711680, _
              Optional ByVal markerStyle As XlMarkerStyle = xlMarkerStyleDiamond, Optional ByVal stringChartTitle As String = "Chart Title")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine creates a new graph from the data based on the row longRowDataHeader
' this is a single series of data NOT multiple series of data
'
' Requirements:
' Cells_ReturnNumberOrLetters()
'
' Inputs:
' wksGraph
' Type: worksheet
' Desc: the worksheet to place the graphs
'
' wksData
' Type: worksheet
' Desc: the worksheet to pull the data
'
' longRowDataHeader
' Type: long
' Desc: the row of the header of the data, below is the sequence of data
' longRowDataHeader + 1 -> date of data collected
' longRowDataHeader + 2 -> data to be displayed
'
' typeGraph
' Type: XlChartType
' Desc: the type of graph to create
'
' longSeriesColor
' Type: long
' Desc: the color for the data series plotted
'
' longColStart
' Type: long
' Desc: the start column for the data to be displayed, the default is column 01
'
' longColStop
' Type: long
' Desc: the stop column for the data to be displayed, the default is column 10
'
' longGraphTopPosit
' Type: long
' Desc: the position of the top of the graph in pixels, default is 1
'
' longGraphLeftPosit
' Type: long
' Desc: the left position of of the graph, default is 1
'
' longGraphWidth
' Type: long
' Desc: the the width of the graph, default is 50
'
' longGraphHeight
' Type: long
' Desc: the height of the graph, default is 50
'
' longMarkerColor
' Type: long
' Desc: the color for the marker if used
'
' markerStyle
' Type: xlMarkerStyle
' Desc: the style of the marker
'
' stringChartTitle
' Type: string
' Desc: The title of the chart
'
' Important Info:
' the default will be 10 dates, if there are less then 10 dates then the graph will display all the graphs
'
' Return:
' None
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' declare variables
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Const boolError As Boolean = False
Dim shapeChart As Shape
Dim chartNewChart As Chart
Dim charttitleNew As ChartTitle
Dim seriesData As Series
Dim axisX As Axis, axisY As Axis
Dim axisXTitle As AxisTitle, axisYTitle As AxisTitle
Dim stringDataSeriesYValues As String, stringDataSeriesXValues As String, stringColStart As String, stringColStop As String
Dim dateSerialStart As Date, dateSerialStop As Date, dateStart As Date, dateStop As Date
 
' loop
Dim a As Long, b As Long, c As Long
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' set objects
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Set shapeChart = wksGraph.Shapes.AddChart(typeGraph, longGraphLeftPosit, longGraphTopPosit, longGraphWidth, longGraphHeight)
Set chartNewChart = shapeChart.Chart
chartNewChart.HasTitle = True
chartNewChart.HasAxis(xlValue) = True
chartNewChart.HasAxis(xlCategory) = True
Set charttitleNew = chartNewChart.ChartTitle
Set axisX = chartNewChart.Axes(xlCategory)
axisX.HasTitle = True
Set axisXTitle = axisX.AxisTitle
Set axisY = chartNewChart.Axes(xlValue)
axisY.HasTitle = True
Set axisYTitle = axisY.AxisTitle
Set seriesData = chartNewChart.SeriesCollection.NewSeries
chartNewChart.HasLegend = False
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' initialize variables
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
 
stringDataSeriesYValues = "tsma"
stringDataSeriesXValues = "tsma"
stringColStart = Cells_ReturnNumberOrLetters(Cells(longRowDataHeader, longColStart), 2)
stringColStop = Cells_ReturnNumberOrLetters(Cells(longRowDataHeader, longColStop), 2)
dateStart = Cells(longRowDataHeader + 1, longColStart).Value
dateStop = Cells(longRowDataHeader + 1, longColStop).Value
dateSerialStart = DateSerial(DatePart("yyyy", dateStart), DatePart("m", dateStart), DatePart("d", dateStart))
dateSerialStop = DateSerial(DatePart("yyyy", dateStop), DatePart("m", dateStop), DatePart("d", dateStop))
 
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
 
' activate destination worksheet
wksGraph.Activate
 
' add data series to chart
stringDataSeriesYValues = "=" & wksData.Name & "!$" & stringColStart & "$" & CStr(longRowDataHeader + 2) & ":$" & stringColStop & "$" & CStr(longRowDataHeader + 2)
stringDataSeriesXValues = "=" & wksData.Name & "!$" & stringColStart & "$" & CStr(longRowDataHeader + 1) & ":$" & stringColStop & "$" & CStr(longRowDataHeader + 1)
 
seriesData.Values = stringDataSeriesYValues
seriesData.XValues = stringDataSeriesXValues
seriesData.Name = Cells(longRowDataHeader, 1).Value
seriesData.Format.Line.ForeColor.RGB = longSeriesColor
seriesData.Format.Line.Weight = 4
 
' test for marker types
Select Case typeGraph
    Case xlLineMarkers, xlLineMarkersStacked, xlLineMarkersStacked100, xlRadarMarkers, xlXYScatter, xlXYScatterLines, xlXYScatterSmooth:
        seriesData.markerStyle = markerStyle
        seriesData.MarkerSize = 20
        seriesData.MarkerForegroundColor = longMarkerColor
        seriesData.MarkerBackgroundColor = longMarkerColor
    Case Else
        seriesData.markerStyle = xlMarkerStyleNone
End Select
 
' add title to chart and axes
charttitleNew.Caption = stringChartTitle
axisYTitle.Caption = "Number of " & wksData.Cells(longRowDataHeader, 1).Value
axisXTitle.Caption = "Date range of " & wksData.Cells(longRowDataHeader, 1).Value
 
' set x-scale
If DateDiff("d", dateStart, dateStop) > 10 Then
    axisX.MinimumScale = dateSerialStart
    axisX.MaximumScale = dateSerialStop
    axisX.MajorUnit = CLng((dateSerialStop - dateSerialStart) / 10)
Else ' do nothing
End If
 
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
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' reset objects
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
 
Set shapeChart = Nothing
Set chartNewChart = Nothing
Set seriesData = Nothing
Set axisX = Nothing
Set axisY = Nothing
Set axisXTitle = Nothing
Set axisYTitle = Nothing
 
End Sub
Function Chart_DeleteAll(ByVal wksWorksheet As Worksheet) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine deletes all the graphs in a worksheet
'
' Requirements:
' None
'
' Inputs:
' wksWorksheet
' Type: worsheet
' Desc: the worksheet to look for the graphs on
'
' Important Info:
' None
'
' Return:
' variable
' Type: boolean
' Desc: if the chart object count is zero will return false / if there are charts and deleted will turn true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' declare variables
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Const boolError As Boolean = False
Dim longChartCount As Long
Dim boolReturnDeletedCharts As Boolean
 
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

longChartCount = wksWorksheet.ChartObjects.Count
boolReturnDeletedCharts = False
 
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
 
' loop through objects on worksheet
If longChartCount > 0 Then
    ' set return value
    boolReturnDeletedCharts = True
 
    ' loop through objects and delete
    For a = 1 To longChartCount
        wksWorksheet.ChartObjects(1).Delete
    Next a
Else ' do nothing
End If
 
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

Chart_DeleteAll = boolReturnDeletedCharts
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'
' reset objects
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
End Function
Sub Chart_Line(wks_source As Worksheet, wks_dest As Worksheet, longNumDataSeries As Long, longChartType As Long, longChartLayout As Long, _
               longDataSeriesStartRow As Long, stringDataSeriesStartColumn As String, stringDataSeriesStopColumn As String, stringGraphTitle As String, _
               stringYAxisLabel As String, stringXAxisLabel As String, longGraphTopPosit As Long, longGraphLeftPosit As Long, _
               longGraphWidth As Long, longGraphHeight As Long)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine creats a line graph with the the information provided
'
' Inputs:
' wks_source
' Type: worksheet object
' Desc: source worksheet
'
' wks_dest
' Type: worksheet object
' Desc: destination worksheet
'
' longNumDataSeries
' Type: long
' Desc: number of data series for the chart
'
' longChartType
' Type: long
' Desc: type of chart to graph
'
' longChartLayout
' Type: long
' Desc: layout of chart
'
' longDataSeriesStartRow
' Type: long
' Desc: the start row of the data series; assume it's the header row
'
' stringDataSeriesStartColumn
' Type: string
' Desc: start column for the data series
'
' stringDataSeriesStopColum
' Type: string
' Desc: stop column for the data series
'
' stringGraphTitle
' Type: string
' Desc: title of graph
'
' stringYAxisLabel
' Type: string
' Desc: "Y" axis label
'
' stringXAxisLabel
' Type: string
' Desc: "X" axis lable
'
' longGraphTopPosit
' Type: long
' Desc: top boundry of graph, in pixels
'
' longGraphLeftPosit
' Type: long
' Desc: left bourndry of graph, in pixels
'
' longGraphWidth
' Type: long
' Desc: width of graph, in pixels
'
' longGraphHeight
' Type: long
' Desc: height of graph, in pixels
'
' Return:
' Type: none
' Desc: none
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' declare
Dim chart_new As Shape
Dim longDataSeriesRow As Long
Dim stringSeriesName As String

' loop variables
Dim a As Long, b As Long, c As Long

' initialize
longDataSeriesRow = 1
stringSeriesName = "tsma"

a = 1

' start
Worksheets(wks_dest).Activate
Range("A1").Select

' create line graph
Set chart_new = wks_dest.Shapes.AddChart
chart_new.ChartType = longChartType

' add graph title and axes
With chart_new
    .HasTitle = True
    .ChartTitle.Text = stringGraphTitle
End With
With chart_new.Axes(xlCategory)
    .HasTitle = True
    .AxisTitle.Text = stringXAxisLabel
End With
With chart_new.Axes(xlValue)
    .HasTitle = True
    .AxisTitle.Text = stringYAxisLabel
End With

' add series data
For a = 1 To longNumDataSeries
    longDataSeriesRow = longDataSeriesStartRow + a
    stringSeriesName = wks_source.Cells(longDataSeriesRow, 1).Value
    chart_new.SeriesCollection.NewSeries
    chart_new.SeriesCollection(a).Name = stringSeriesName
    chart_new.SeriesCollection(a).Values = "='" & wks_source.Name & "'!$" & stringDataSeriesStartColumn & "$" & longDataSeriesRow & _
                                             ":$" & stringDataSeriesStopColumn & "$" & longDataSeriesRow
    chart_new.SeriesCollection(a).XValues = "='" & wks_source.Name & "'!$" & stringDataSeriesStartColumn & "$" & longDataSeriesStartRow & _
                                              ":$" & stringDataSeriesStopColumn & "$" & longDataSeriesStartRow
Next a

' sets chart position and size
chart_new.Left = longGraphLeftPosit
chart_new.Top = longGraphTopPosit
chart_new.Width = longGraphWidth
chart_new.Height = longGraphHeight

' apply layout
chart_new.ApplyLayout (longChartLayout)

' object cleanup
End Sub
