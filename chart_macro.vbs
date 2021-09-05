Sub chart_data(inputRange As String, outputRange As String)
'
' chart_data Subroutine
' Charts the data represented by inputRange and outputRange.
'
    Dim totalRange As String    ' Complete range to graph
    
    ' Set up the range
    totalRange = Join(Array(inputRange, outputRange), ",")
    
    ' Select the range
    Range(totalRange).Select
    
    MsgBox "Hello, world!"
    
End Sub 'chart_data(inputRange As String, outputRange As String, titleRow As Integer)


