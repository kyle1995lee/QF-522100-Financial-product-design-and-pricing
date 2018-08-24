Attribute VB_Name = "Module2"
Function BSOptionValue(iopt, S, X, r, q, tyr, sigma)
    Dim eqt, ert, NDOne, NDTwo
    eqt = Exp(-q * tyr)
    ert = Exp(-r * tyr)
    If S > 0 And X > 0 And tyr > 0 And sigma > 0 Then
        NDOne = Application.NormSDist(iopt * BSDOne(S, X, r, q, tyr, sigma))
        NDTwo = Application.NormSDist(iopt * BSDTwo(S, X, r, q, tyr, sigma))
        BSOptionValue = iopt * (S * eqt * NDOne - X * ert * NDTwo)
    Else
        BSOptionValue = -1
    End If
    
End Function

Function BSDOne(S, X, r, q, tyr, sigma)
    BSDOne = (Log(S / X) + (r - q + 0.5 * sigma ^ 2) * tyr) / (sigma * Sqr(tyr))
End Function

Function BSDTwo(S, X, r, q, tyr, sigma)
    BSDTwo = (Log(S / X) + (r - q - 0.5 * sigma ^ 2) * tyr) / (sigma * Sqr(tyr))
End Function

Function Chooser(S, X, r, q, ct, tyr, sigma)
    Dim Xp
    Xp = X * Exp(-1 * (r - q) * (tyr - ct))
    If S > 0 And X > 0 And tyr > 0 And sigma > 0 Then
        Chooser = BSOptionValue(1, S, X, r, q, tyr, sigma) + Exp(-q * (tyr - ct)) * BSOptionValue(-1, S, Xp, r, q, ct, sigma)
    Else
       Chooser = -1
    End If
End Function



Private Sub OptionButton1_Click()
    Range("A19:Z58").Select
    Selection.ClearContents
    Worksheets("Package").Activate
    For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3) = ST
    Next S
    For t = 1 To 4
      Tt = (t - 1) * 1 + 0.0001
      For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3 + t) = BSOptionValue(1, ST, Cells(5, 2), Cells(6, 2), Cells(8, 2), Tt, Cells(12, 2)) _
                             - BSOptionValue(1, ST, Cells(5, 3), Cells(6, 2), Cells(8, 2), Tt, Cells(12, 2))
      Next S
    Next t
    
    ActiveSheet.ChartObjects("圖表 1").Activate
    ActiveChart.Parent.Delete
   
    Charts.Add
    ActiveChart.ChartType = xlLine
    ActiveChart.SetSourceData Source:=Sheets("Package").Range("C18:G28"), _
        PlotBy:=xlColumns
    ActiveChart.SeriesCollection(1).Delete
    ActiveChart.SeriesCollection(1).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(2).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(3).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(4).XValues = "=Package!R19C3:R28C3"
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Package"
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "Bull Spread Call"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "S"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Payoff"
    End With


End Sub

Private Sub OptionButton2_Click()
    Selection.ClearContents
    Worksheets("Package").Activate
    For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3) = ST
    Next S
    For t = 1 To 4
      Tt = (t - 1) * 1 + 0.0001
      For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3 + t) = BSOptionValue(-1, ST, Cells(5, 2), Cells(6, 2), Cells(8, 2), Tt, Cells(12, 2)) _
                             - BSOptionValue(-1, ST, Cells(5, 3), Cells(6, 2), Cells(8, 2), Tt, Cells(12, 2))
      Next S
    Next t
    
    ActiveSheet.ChartObjects("圖表 1").Activate
    ActiveChart.Parent.Delete
    
    Charts.Add
    ActiveChart.ChartType = xlLine
    ActiveChart.SetSourceData Source:=Sheets("Package").Range("C18:G28"), _
        PlotBy:=xlColumns
    ActiveChart.SeriesCollection(1).Delete
    ActiveChart.SeriesCollection(1).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(2).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(3).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(4).XValues = "=Package!R19C3:R28C3"
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Package"
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "Bull Spread Put"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "S"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Payoff"
    End With


End Sub

Private Sub OptionButton3_Click()
    Selection.ClearContents
    Worksheets("Package").Activate
    For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3) = ST
    Next S
    For t = 1 To 4
      Tt = (t - 1) * 1 + 0.0001
      For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3 + t) = BSOptionValue(1, ST, Cells(5, 3), Cells(6, 2), Cells(8, 2), Tt, Cells(12, 2)) _
                             - BSOptionValue(1, ST, Cells(5, 2), Cells(6, 2), Cells(8, 2), Tt, Cells(12, 2))
      Next S
    Next t
    
    ActiveSheet.ChartObjects("圖表 1").Activate
    ActiveChart.Parent.Delete
    
    Charts.Add
    ActiveChart.ChartType = xlLine
    ActiveChart.SetSourceData Source:=Sheets("Package").Range("C18:G28"), _
        PlotBy:=xlColumns
    ActiveChart.SeriesCollection(1).Delete
    ActiveChart.SeriesCollection(1).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(2).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(3).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(4).XValues = "=Package!R19C3:R28C3"
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Package"
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "Bear Spread Call"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "S"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Payoff"
    End With

End Sub

Private Sub OptionButton4_Click()
    Selection.ClearContents
    Worksheets("Package").Activate
    For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3) = ST
    Next S
    For t = 1 To 4
      Tt = (t - 1) * 1 + 0.0001
      For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3 + t) = BSOptionValue(-1, ST, Cells(5, 3), Cells(6, 2), Cells(8, 2), Tt, Cells(12, 2)) _
                             - BSOptionValue(-1, ST, Cells(5, 2), Cells(6, 2), Cells(8, 2), Tt, Cells(12, 2))
      Next S
    Next t
    
    ActiveSheet.ChartObjects("圖表 1").Activate
    ActiveChart.Parent.Delete
    
    Charts.Add
    ActiveChart.ChartType = xlLine
    ActiveChart.SetSourceData Source:=Sheets("Package").Range("C18:G28"), _
        PlotBy:=xlColumns
    ActiveChart.SeriesCollection(1).Delete
    ActiveChart.SeriesCollection(1).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(2).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(3).XValues = "=Package!R19C3:R28C3"
    ActiveChart.SeriesCollection(4).XValues = "=Package!R19C3:R28C3"
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Package"
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "Bear Spread Put"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "S"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Payoff"
    End With

End Sub


