Sub chart_data(inputRange As String, outputRange As String)
'
' chart_data Subroutine
' Charts the data represented by inputRange and outputRange.
' by      : Leomar Duran
' selft   : https://github.com/lduran2/excel-chart-macro
' created : 2021-09-05 t01:52
' version : 1.1
'
' changelog :
'     v1.1 -- 2021-09-05 t01:57
'         creates the chart
'
'     v1.0 -- 2021-09-05 t01:52
'         selects the data and prints a greeting
'
    Dim totalRange As String    ' Complete range to graph

    ' Set up the range
    totalRange = Join(Array(inputRange, outputRange), ",")

    ' Select the range
    Range(totalRange).Select

    ' Create the scatter plot chart from this range
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines).Select

    MsgBox "Hello, world!"

End Sub 'chart_data(inputRange As String, outputRange As String)


