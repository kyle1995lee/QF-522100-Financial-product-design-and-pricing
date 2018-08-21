Attribute VB_Name = "ExporUse"
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
Function CashorNothing(iopt, S, X, K, r, q, tyr, sigma)
'   Returns Black-Scholes Value (iopt=1 for call, -1 for put; q=div yld)
'   Uses BSDOne fn
'   Uses BSDTwo fn
    Dim eqt, ert, NDOne, NDTwo
    eqt = Exp(-q * tyr)
    ert = Exp(-r * tyr)
    If S > 0 And X > 0 And tyr > 0 And sigma > 0 Then
        NDOne = Application.NormSDist(iopt * BSDOne(S, X, r, q, tyr, sigma))
        NDTwo = Application.NormSDist(iopt * BSDTwo(S, X, r, q, tyr, sigma))
        CashorNothing = K * ert * NDTwo
    Else
       CashorNothing = -1
    End If
End Function

Function AssetorNothing(iopt, S, X, K, r, q, tyr, sigma)
'   Returns Black-Scholes Value (iopt=1 for call, -1 for put; q=div yld)
'   Uses BSDOne fn
'   Uses BSDTwo fn
    Dim eqt, ert, NDOne, NDTwo
    eqt = Exp(-q * tyr)
    ert = Exp(-r * tyr)
    If S > 0 And X > 0 And tyr > 0 And sigma > 0 Then
        NDOne = Application.NormSDist(iopt * BSDOne(S, X, r, q, tyr, sigma))
        NDTwo = Application.NormSDist(iopt * BSDTwo(S, X, r, q, tyr, sigma))
        AssetorNothing = S * eqt * NDOne
    Else
       AssetorNothing = -1
    End If
End Function


Sub Gohome()
    Sheets("Home").Select
End Sub
Sub CommandButton1_Click()
    Sheets("Excel").Select
End Sub

Sub CommandButton2_Click()
    Sheets("Vba_Function").Select
End Sub

Sub CommandButton3_Click()
    Sheets("Vba_Sub").Select
End Sub

Sub CommandButton4_Click()
    Sheets("Chart").Select
End Sub

Sub CommandButton5_Click()
    Sheets("Chart_Table").Select
End Sub

Sub CommandButton6_Click()
    Sheets("Chart_Sub").Select
End Sub

Sub CommandButton7_Click()
    Sheets("Binary_Option").Select
End Sub

Sub BSOptionCalc()
   Worksheets("vba_sub").Activate
   Cells(5, 5) = BSOptionValue(1, Cells(4, 2), Cells(5, 2), Cells(6, 2), Cells(8, 2), Cells(11, 2), Cells(12, 2))
   Cells(5, 8) = BSOptionValue(-1, Cells(4, 2), Cells(5, 2), Cells(6, 2), Cells(8, 2), Cells(11, 2), Cells(12, 2))
End Sub

Sub OptionButton1_Click()
    Range("A19:Z58").Select
    Selection.ClearContents
    Worksheets("chart_sub").Activate
    For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3) = ST
    Next S
    For T = 1 To 4
      Tt = (T - 1) * 1 + 0.0001
      For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3 + T) = BSOptionValue(-1, ST, Cells(5, 2), Cells(6, 2), Cells(8, 2), Tt, Cells(12, 2))
      Next S
    Next T
    
    ActiveSheet.ChartObjects.Delete
    
    Charts.Add
    ActiveChart.ChartType = xlLine
    ActiveChart.SetSourceData Source:=Sheets("Chart_Sub").Range("C18:G28"), _
        PlotBy:=xlColumns
    ActiveChart.SeriesCollection(1).Delete
    ActiveChart.SeriesCollection(1).XValues = "=chart_sub!R19C3:R28C3"
    ActiveChart.SeriesCollection(2).XValues = "=chart_sub!R19C3:R28C3"
    ActiveChart.SeriesCollection(3).XValues = "=chart_sub!R19C3:R28C3"
    ActiveChart.SeriesCollection(4).XValues = "=chart_sub!R19C3:R28C3"
    ActiveChart.Location Where:=xlLocationAsObject, Name:="chart_sub"
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "Put"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "S"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Payoff"
    End With
End Sub
Sub OptionButton2_Click()
    Range("A19:Z58").Select
    Selection.ClearContents
    Worksheets("chart_sub").Activate
    For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3) = ST
    Next S
    For T = 1 To 4
      Tt = (T - 1) * 1 + 0.0001
      For S = 1 To 10
        ST = 20 * S
        Cells(18 + S, 3 + T) = BSOptionValue(1, ST, Cells(5, 2), Cells(6, 2), Cells(8, 2), Tt, Cells(12, 2))
      Next S
    Next T
    
    ActiveSheet.ChartObjects.Delete
    
    Charts.Add
    ActiveChart.ChartType = xlLine
    ActiveChart.SetSourceData Source:=Sheets("Chart_Sub").Range("C18:G28"), _
        PlotBy:=xlColumns
    ActiveChart.SeriesCollection(1).Delete
    ActiveChart.SeriesCollection(1).XValues = "=chart_sub!R19C3:R28C3"
    ActiveChart.SeriesCollection(2).XValues = "=chart_sub!R19C3:R28C3"
    ActiveChart.SeriesCollection(3).XValues = "=chart_sub!R19C3:R28C3"
    ActiveChart.SeriesCollection(4).XValues = "=chart_sub!R19C3:R28C3"
    ActiveChart.Location Where:=xlLocationAsObject, Name:="chart_sub"
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "Call"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "S"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Payoff"
    End With
End Sub

Sub CashorNothingCall()
    If Range("B4") >= Range("B5") Then
        Range("D6") = CashorNothing(1, Range("B4"), Range("B5"), Range("B6"), Range("B7"), Range("B9"), Range("B12"), Range("B13"))
    Else
        Range("D6") = 0
    End If
End Sub

Sub CashorNothingPut()
    If Range("B4") >= Range("B5") Then
        Range("D6") = CashorNothing(-1, Range("B4"), Range("B5"), Range("B6"), Range("B7"), Range("B9"), Range("B12"), Range("B13"))
    Else
        Range("D6") = 0
    End If
End Sub

Sub AssetorNothingCall()
    If Range("B4") >= Range("B5") Then
        Range("D6") = AssetorNothing(1, Range("B4"), Range("B5"), Range("B6"), Range("B7"), Range("B9"), Range("B12"), Range("B13"))
    Else
        Range("D6") = 0
    End If
End Sub

Sub AssetorNothingPut()
    If Range("B4") >= Range("B5") Then
        Range("D6") = AssetorNothing(-1, Range("B4"), Range("B5"), Range("B6"), Range("B7"), Range("B9"), Range("B12"), Range("B13"))
    Else
        Range("D6") = 0
    End If
End Sub


Sub black_scholes()
    UserForm1.Show
End Sub

