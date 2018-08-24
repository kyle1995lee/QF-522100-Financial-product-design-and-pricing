Attribute VB_Name = "Module1"

Function BinOptionValueTest(iopt, iea, S, r, tyr, sigma, nstep, cr, f, ct, pt)
'   Returns Binomial Option Value (imod=1 for CRR, 2 for LR; iea=1 for euro, 2 for amer)
'   Uses BSDOne fn
'   Uses BSDTwo fn
'   Uses PPNormInv fn
    Dim delt, erdt, ermqdt, u, d, d1, d2, p, pdash, pstar
    Dim i As Integer, j As Integer
    Dim vvec() As Variant
    If imod = 2 Then nstep = Application.Odd(nstep)
    ReDim vvec(nstep)
    
If S > 0 And tyr > 0 And sigma > 0 And f > 0 Then
        delt = tyr / nstep
        erdt = Exp(r * delt)
        ermqdt = Exp((r - q) * delt)
        u = Exp(sigma * Sqr(delt))
        d = 1 / u
        p = (ermqdt - d) / (u - d)
        pstar = 1 - p

    
    For i = 0 To nstep
      vvec(i) = Application.Max(iopt * (S * (u ^ i) * (d ^ (nstep - i)) * cr), f)
    Next i
    
    For j = nstep - 1 To 0 Step -1
        For i = 0 To j
            vvec(i) = (p * vvec(i + 1) + pstar * vvec(i)) / erdt
            If iea = 2 Then
                vvec(i) = Application.Max(vvec(i), iopt * (S * (u ^ i) * (d ^ (j - i))) * cr)
            
            ElseIf iea = 3 Then
                If j > nstep / 2 Then
                    vvec(i) = Application.Max(Application.Min(vvec(i), ct), iopt * (S * (u ^ i) * (d ^ (j - i))) * cr)
                Else
                    vvec(i) = Application.Max(vvec(i), iopt * (S * (u ^ i) * (d ^ (j - i))) * cr)
                End If
            
            ElseIf iea = 4 Then
                If j = nstep / 2 Then
                    vvec(i) = Application.Max(Application.Min(Application.Max(vvec(i), pt), ct), iopt * (S * (u ^ i) * (d ^ (j - i)) * cr))
                ElseIf j > nstep / 2 Then
                    vvec(i) = Application.Max(Application.Min(vvec(i), ct), iopt * (S * (u ^ i) * (d ^ (j - i))) * cr)
                Else
                    vvec(i) = Application.Max(vvec(i), iopt * (S * (u ^ i) * (d ^ (j - i))) * cr)
                End If
            
            End If
        Next i
    Next j
    BinOptionValueTest = vvec(0)
    Else
    BinOptionValueTest = -1
    End If
End Function


Function PPNormInv(z, n)
'   Returns the Peizer-Pratt Inversion
'   Only defined for n odd
'   Used in LR Binomial Option Valuation
    Dim c1
    n = Application.Odd(n)
    c1 = Exp(-((z / (n + 1 / 3 + 0.1 / (n + 1))) ^ 2) * (n + 1 / 6))
    PPNormInv = 0.5 + Sgn(z) * Sqr((0.25 * (1 - c1)))
End Function
Function BinTree(imod, S, r, q, tyr, sigma, nstep)
'   Returns Binomial Share Price Tree (imod=0 for JR, 1 for CRR)
    Dim delt, rnmut, u, d
    Dim i As Integer, j As Integer
    Dim Smat() As Variant
    ReDim Smat(nstep, nstep)
    delt = tyr / nstep
    If imod = 0 Then
       rnmut = (r - q - 0.5 * sigma ^ 2) * delt
       u = Exp(rnmut + sigma * Sqr(delt))
       d = Exp(rnmut - sigma * Sqr(delt))
    Else
       u = Exp(sigma * Sqr(delt))
       d = 1 / u
    End If
     Smat(nstep, 0) = S
     
    For i = 1 To nstep
        Smat(nstep - i, 0) = ""
    Next i
    For j = 1 To nstep
        For i = 0 To j - 1
            Smat(nstep - i, j) = d * Smat(nstep - i, j - 1)
        Next i
            Smat(nstep - j, j) = u * Smat(nstep - j + 1, j - 1)
        For i = j + 1 To nstep
            Smat(nstep - i, j) = ""
        Next i
    Next j
    BinTree = Smat
End Function
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

Sub MCIntegration()
    Dim meanexp
    Dim i As Integer
    Randomize
    Number = Cells(3, 4)
    b = Cells(4, 4)
    a = Cells(5, 4)
    meanexp = 0
    For i = 1 To Number
       meanexp = Exp(Rnd * (b - a) + a) + meanexp
    Next i
    Cells(4, 2) = meanexp * Sqr((b - a) ^ 2) / Number
End Sub


Function MCOptionValue(iopt, S, X, r, q, tyr, sigma, nsim)

'   Simple-ish Monte-Carlo simulation to value BS Option
'   Uses antithetic variate (-rands)

    Dim rnmut, sigt, sum, randns, S1, S2, payoff1, payoff2
    Dim i As Integer
    Randomize
    rnmut = (r - q - 0.5 * sigma ^ 2) * tyr
    sigt = sigma * Sqr(tyr)
    sum = 0
   For i = 1 To nsim
        randns = Application.NormSInv(Rnd)
        S1 = S * Exp(rnmut + randns * sigt)
        S2 = S * Exp(rnmut - randns * sigt)
        payoff1 = Application.Max(iopt * (S1 - X), 0)
        payoff2 = Application.Max(iopt * (S2 - X), 0)
        sum = sum + 0.5 * (payoff1 + payoff2)
    Next i

    MCOptionValue = Exp(-r * tyr) * sum / nsim
End Function

Sub shareprice()
    Dim rnmut, sigt, randns, S
    Dim i As Integer
    Randomize
    Range("A19:z1000").Select
    Selection.ClearContents

   S = Cells(4, 2)
    r = Cells(6, 2)
    q = Cells(8, 2)
    tyr = Cells(11, 2)
    sigma = Cells(12, 2)
    nsim = Cells(14, 2)
    rnmut = (r - q - 0.5 * sigma ^ 2) * (tyr / nsim)
    sigt = sigma * Sqr(tyr / nsim)
    Cells(21, 1) = Cells(14, 1)
    Cells(21, 5) = Cells(4, 1)
    Cells(22, 1) = 1
    Cells(22, 5) = Cells(4, 2)
    For i = 2 To nsim
        randns = Application.NormSInv(Rnd)
        Cells(21 + i, 5) = Cells(21 + i - 1, 5) * Exp(rnmut + randns * sigt)
        Cells(21 + i, 1) = i
    Next i
    ActiveSheet.ChartObjects.Delete
    Range("E22:E121").Select
    
    ActiveSheet.Shapes.AddChart2(227, xlLineMarkers).Select
    ActiveChart.SetSourceData Source:=Range("Share_Price!$E$22:$E$121")


End Sub





