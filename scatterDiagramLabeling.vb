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
      'See the 1st row of the sheet above the x values in the sheet.
      ActiveChart.SeriesCollection(Series).Points(Counter).DataLabel.Text = _
        Cells(1, Range(xVals).Cells(Counter).Column).Value
        
    Next Counter
  Next Series

End Sub


'グラフの各要素に対し、直上のセルの（その要素のx値があるセル(x,y)に対し、(1,y)のセルの）値をラベルを加えるマクロ
'対象のグラフをアクティブにしてからマクロを実行することで利用できる。

'シートの例（空行挿入可。x,yは逆順でも可）
'label1, label2, label3,...
'x11,x12,x13,...
'y11,y12,y13,...
'x21,x22,x23,...
'y21,y22,y23,...
'...
