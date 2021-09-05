Sub chart_cols(inputColumn As String, outputColumn As String, _
    titleRow As Integer _
)
'
' chart_cols Subroutine
' Charts the data represented by columns inputColumn and outputColumn
' for convenience.
'
' by      : Leomar Duran <https://github.com/lduran2/>
' when    : 2021-09-05 t01:52
' self    : https://github.com/lduran2/excel-chart-macro
' version : 1.2
'
' changelog :
'     v1.2 -- 2021-09-05 t02:46
'         added chart title, abstracted cell coordinating
'
'     v1.1 -- 2021-09-05 t01:57
'         created the chart
'
'     v1.0 -- 2021-09-05 t01:52
'         selectd the data and printed a greeting
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
    chart_data totalRange, inputTitleCell, outputTitleCell

End Sub 'chart_cols(inputColumn As String, outputColumn As String, _
'   titleRow As Integer _
' )
' --------------------------------------------------------------------

Sub chart_data(dataRange As String, _
    inputTitleCell As String, outputTitleCell As String _
)
'
' chart_data Subroutine
' Charts the data represented by dataRange.
'
    Dim inputTitle As String    ' title of input data
    Dim outputTitle As String   ' title of output data
    Dim chartTitle As String    ' title of the chart
    
    ' Get the input and output titles
    inputTitle = Range(inputTitleCell).Value
    outputTitle = Range(outputTitleCell).Value
    
    ' Make the chart title
    chartTitle = Join(Array(inputTitle, outputTitle), " vs ")
    
    ' Select the range
    Range(dataRange).Select

    ' Create the scatter plot chart from this range
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines).Select
    
    ' Add the chart title text
    ActiveChart.chartTitle.Text = chartTitle

    MsgBox "Hello, world!"

End Sub 'Sub chart_data(dataRange As String, _
'   inputTitleCell As String, outputTitleCell As String _
' )
