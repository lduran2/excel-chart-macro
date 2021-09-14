Sub chart_cols(summary_sheet As worksheet, supertitle As Variant, _
    inputColumn As String, outputColumn As Variant, _
    outputBounds As Variant, _
    titleRow As Integer, _
    chartDim As Dimension _
)
' ./chart_macro.vbs
' chart_cols Subroutine
' Plots the chart of the data represented by columns inputColumn and
' outputColumn for convenience.
'
' by      : Leomar Duran <https://github.com/lduran2/>
' when    : 2021-09-14 t05:48
' self    : https://github.com/lduran2/excel-chart-macro
' version : 2.1
'
' changelog :
'     v2.1 -- 2021-09-14 t05:48
'         repositioned and resized charts in the "Chart Summary" worksheet
'
'     v2.0 -- 2021-09-14 t04:30
'         integrated with interface
'         plots charts from the copied worksheet data on the
'             "Chart Summary" worksheet
'
'     v1.4 -- 2021-09-05 t03:29
'         styled the chart's series line
'
'     v1.3.1 -- 2021-09-05 t03:10
'         abstracted axes title labels
'
'     v1.3 -- 2021-09-05 t03:02
'         added axes titles
'
'     v1.2 -- 2021-09-05 t02:46
'         added chart title, abstracted cell coordinating
'
'     v1.1 -- 2021-09-05 t01:57
'         created the chart
'
'     v1.0 -- 2021-09-05 t01:52
'         selectd the data and printed a greeting
'
'
    Dim inputRange As String    ' Range to select for input
    Dim outputRange As String   ' Range to select for output
    Dim totalRange As String    ' Complete range to graph

    Dim inputTitleCell As String    ' Cell for title of the input data
    Dim outputTitleCell As String   ' Cell for title of the output data

    ' Set up input and output range
    inputRange = Join(Array(inputColumn, inputColumn), ":")
    outputRange = Join(Array(outputColumn, outputColumn), ":")
    ' Set up data range
    totalRange = Join(Array(inputRange, outputRange), ",")

    ' Set up title cells
    inputTitleCell = Join(Array(inputColumn, CStr(titleRow)), "")
    outputTitleCell = Join(Array(outputColumn, CStr(titleRow)), "")

    ' Delegate to chart_data
    chart_data summary_sheet, supertitle, _
        totalRange, inputTitleCell, outputTitleCell, outputBounds, _
        chartDim

End Sub 'chart_cols(summary_sheet As worksheet, supertitle As Variant, _
'   inputColumn As String, outputColumn As Variant, _
'   outputBounds As Variant,
'   titleRow As Integer, _
    chartDim As Dimension _
' )
' --------------------------------------------------------------------

Sub chart_data(summary_sheet As worksheet, supertitle As Variant, _
    dataRange As String, _
    inputTitleCell As String, outputTitleCell As String, _
    outputBounds As Variant, _
    chartDim As Dimension _
)
'
' chart_data Subroutine
' Plots the chart of the data represented by dataRange.
'
    Const chartStyle = 240      ' The style of the chart, corresponds to xlXYScatterLines

    Dim inputTitle As String    ' title of input data
    Dim outputTitle As String   ' title of output data
    Dim chartTitle As String    ' title of the chart
    Dim dataSheet As worksheet  ' this worksheet which the data comes from

    ' Store the data sheet
    Set dataSheet = ActiveSheet

    ' Get the input and output titles
    inputTitle = Range(inputTitleCell).value
    outputTitle = Range(outputTitleCell).value

    ' Make the chart title
    chartTitle = Join(Array(supertitle, ": ", outputTitle, " vs ", inputTitle), "")

    ' Select the range
    Range(dataRange).Select

    ' Create the scatter plot chart from this range
    ActiveSheet.Shapes.AddChart2(chartStyle, xlXYScatterLines).Select

    ' Add the chart title text
    ActiveChart.chartTitle.Text = chartTitle
    ' Add the transpose axis label
    label_axes ActiveChart, xlCategory, inputTitle
    ' Add the vertical axis label
    label_axes ActiveChart, xlValue, outputTitle

    ' Bound the vertical axis
    bound_axis ActiveChart.Axes(xlValue, xlPrimary), outputBounds

    ' Style the series line
    style_chart_series ActiveChart.FullSeriesCollection(1)
    
    ' Move the chart to the summary sheet
    ActiveChart.Parent.Cut
    summary_sheet.Select
    ActiveSheet.Paste
    
    ' Reposition and resize
    repose_resize ActiveSheet.Shapes(ActiveSheet.Shapes.Count), chartDim
    
    ' Go back to the data
    dataSheet.Activate

End Sub 'chart_data(summary_sheet As worksheet, supertitle As Variant, _
'   dataRange As String, _
'   inputTitleCell As String, outputTitleCell As String, _
    outputBounds As Variant, _
    chartDim As Dimension _
' ) As Chart
' --------------------------------------------------------------------

Sub label_axes(aChart As chart, axesType As Variant, title As String)
'
' label_axes Subroutine
' Labels the axis specified by axesType.
'
    aChart.Axes(axesType, xlPrimary).HasTitle = True
    aChart.Axes(axesType, xlPrimary).AxisTitle.Text = title

End Sub 'label_axes(aChart As chart, axesType As Variant, title As String)
' --------------------------------------------------------------------

Sub style_chart_series(aSeries As Series)
'
' style_chart_series Subroutine
' Styles the chart's series as a thin sky blue line with no markers.
'
    Const MARKER_NONE = -4142   ' display no markers

    ' Remove series markers
    aSeries.MarkerStyle = MARKER_NONE

    ' line thickness [pt]
    aSeries.Format.Line.Weight = 1#
    ' line color [RGB]
    aSeries.Format.Line.ForeColor.RGB = RGB(68, 114, 196)

End Sub 'style_chart_series(aSeries As Series)
' --------------------------------------------------------------------

Sub bound_axis(an_axis As Axis, bounds As Variant)
'
' bound_axis Subroutine
' Bounds a chart axis to bounds [min, max]
'
    an_axis.MinimumScale = bounds(0)
    an_axis.MaximumScale = bounds(1)

End Sub 'bound_axis(an_axis As Axis, bounds As Variant)

Sub repose_resize(a_shape As Shape, a_dimension As Dimension)
'
' bound_axis Subroutine
' Bounds a chart axis to bounds [min, max]
'
    a_shape.IncrementTop (a_dimension.Top - a_shape.Top)
    a_shape.IncrementLeft (a_dimension.Left - a_shape.Left)
    a_shape.ScaleWidth (a_dimension.Width / a_shape.Width), msoFalse, _
        msoScaleFromTopLeft
    a_shape.ScaleHeight (a_dimension.Height / a_shape.Height), msoFalse, _
        msoScaleFromTopLeft

End Sub 'repose_resize(a_shape As Shape, a_dimension As Dimension)


