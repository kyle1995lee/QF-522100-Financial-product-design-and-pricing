Attribute VB_Name = "Module1"
Function Halton1(n, b) As Double
  Dim h As Double, f As Double
  Dim n1 As Integer, n0 As Integer, r As Integer
  n0 = n
  h = 0
  f = 1 / b
  Do While n0 > 0
     n1 = Int(n0 / b)
     r = n0 - n1 * b
     h = h + f * r
     f = f / b
     n0 = n1
  Loop
  Halton1 = h
End Function


Sub Halton_Calc()
    h1 = Cells(3, 2)
    h2 = Cells(3, 3)
    Start = Cells(1, 3)
    n = Cells(3, 1)
    graphtype = Cells(9, 15)
    graphprocess = Cells(11, 15)
    Dim a As Integer
    
    'clear all
    Range("A4:M250").Select
    Selection.ClearContents
    
    'make num list
    For i = 0 To n
        Cells(4 + i, 1) = Start + i
    Next i
    'calc halton1 list
    For i = 0 To n
        Cells(4 + i, 2) = Halton1(Cells(4 + i, 1), h1)
        Cells(4 + i, 3) = Halton1(Cells(4 + i, 1), h2)
        Cells(4 + i, 4) = Rnd
        Cells(4 + i, 5) = Rnd
        'normsinv
        Cells(4 + i, 6) = Application.NormSInv(Cells(4 + i, 2))
        Cells(4 + i, 7) = Application.NormSInv(Cells(4 + i, 3))
        Cells(4 + i, 8) = Application.NormSInv(Cells(4 + i, 4))
        Cells(4 + i, 9) = Application.NormSInv(Cells(4 + i, 5))
        'boxmuller
        Cells(4 + i, 10) = BoxMullerNormSInv1(Cells(4 + i, 2), Cells(4 + i, 3))
        Cells(4 + i, 11) = BoxMullerNormSInv2(Cells(4 + i, 2), Cells(4 + i, 3))
        Cells(4 + i, 12) = BoxMullerNormSInv1(Cells(4 + i, 4), Cells(4 + i, 5))
        Cells(4 + i, 13) = BoxMullerNormSInv2(Cells(4 + i, 4), Cells(4 + i, 5))
    Next i
    
    'clear all sheets
    ActiveSheet.ChartObjects.Delete
    Range("E22:E121").Select

    
    
    a = graphtype * 2 + graphprocess
    a = (a - 2) * 2
    
    Range(Cells(4, a), Cells(4 + n, a + 1)).Select
    
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.SetSourceData Source:=Sheets("Halton_sub").Range(Cells(4, a), Cells(4 + n, a + 1))
    

End Sub
Function BoxMullerNormSInv1(phix1 As Double, phix2 As Double) As Double

'   Replaces NormSInv for quasi-random sequences (eg Faure)
'   See Box and Muller

    Dim h1, h2, vlog, norm1
    h1 = phix1
    h2 = phix2
    vlog = Sqr(-2 * Log(h1))
    norm1 = vlog * Cos(2 * Application.Pi() * h2)
    BoxMullerNormSInv1 = norm1

End Function
Function BoxMullerNormSInv2(phix1 As Double, phix2 As Double) As Double

'   Replaces NormSInv for quasi-random sequences (eg Faure)
'   See Box and Muller

    Dim h1, h2, vlog, norm2
    h1 = phix1
    h2 = phix2
    vlog = Sqr(-2 * Log(h1))
    norm2 = vlog * Sin(2 * Application.Pi() * h2)
    BoxMullerNormSInv2 = norm2

End Function



