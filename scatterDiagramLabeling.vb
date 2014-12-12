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

