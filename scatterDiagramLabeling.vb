'From http://support.microsoft.com/kb/914813/ja
Sub AttachLabelsToPoints()

  'Dimension variables.
  Dim Counter As Integer, ChartName As String, xVals As String

  ' Disable screen updating while the subroutine is run.
  Application.ScreenUpdating = False

  For Series = 1 To ActiveChart.SeriesCollection.Count
   
    'Store the formula for the first series in "xVals".
    xVals = ActiveChart.SeriesCollection(Series).Formula

    'Extract the range for the data from xVals.
    xVals = Mid(xVals, InStr(InStr(xVals, ","), xVals, _
       Mid(Left(xVals, InStr(xVals, "!") - 1), 9)))
    xVals = Left(xVals, InStr(InStr(xVals, "!"), xVals, ",") - 1)
    Do While Left(xVals, 1) = ","
       xVals = Mid(xVals, 2)
    Loop

    'Attach a label to each data point in the chart.
    For Counter = 1 To Range(xVals).Cells.Count
        ActiveChart.SeriesCollection(Series).Points(Counter).HasDataLabel = _
          True
      ActiveChart.SeriesCollection(Series).Points(Counter).DataLabel.Text = _
        Range(xVals).Cells(Counter).Offset(-2 * Series, 0).Value
    Next Counter
  Next Series

End Sub

' label for element 1: la1
' x for series 1 and element 1:x11 
' y for series 1 and element 1:y11 

' Sheet structrue assumption
' la1,la2,la3,...
' x11,x12,x13,...
' y11,y12,y13,...
' x21,x22,x23,...
' y21,y22,y23,...

