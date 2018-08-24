Attribute VB_Name = "Module3"
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


Function BS(S, K, r, d, t, v, cp)
      d1 = (Log(S / K) + (r - d + v ^ 2 / 2) * t) / (v * t ^ 0.5)
      d2 = d1 - v * t ^ 0.5
      BS = cp * S * Exp(-d * t) * Application.NormSDist(cp * d1) - cp * K * Exp(-r * t) * Application.NormSDist(cp * d2)
End Function


Function BSDThree(S, X, r, q, tyr, sigma, Sb)
BSDThree = (Log(S / Sb) + (r - q + 0.5 * sigma ^ 2) * tyr) / (sigma * Sqr(tyr))
End Function
Function BSDFour(S, X, r, q, tyr, sigma, Sb)
BSDFour = (Log(S / Sb) + (r - q + 0.5 * sigma ^ 2) * tyr) / (sigma * Sqr(tyr)) - (sigma * Sqr(tyr))
End Function
Function BSDFive(S, X, r, q, tyr, sigma, Sb)
BSDFive = (Log(S / Sb) - (r - q - 0.5 * sigma ^ 2) * tyr) / (sigma * Sqr(tyr))
End Function
Function BSDSix(S, X, r, q, tyr, sigma, Sb)
BSDSix = (Log(S / Sb) - (r - q - 0.5 * sigma ^ 2) * tyr) / (sigma * Sqr(tyr)) - (sigma * Sqr(tyr))
End Function
Function BSDSeven(S, X, r, q, tyr, sigma, Sb)
BSDSeven = (Log(S * X / Sb ^ 2) - (r - q - 0.5 * sigma ^ 2) * tyr) / (sigma * Sqr(tyr))
End Function
Function BSDEight(S, X, r, q, tyr, sigma, Sb)
BSDEight = (Log(S * X / Sb ^ 2) - (r - q - 0.5 * sigma ^ 2) * tyr) / (sigma * Sqr(tyr)) - (sigma * Sqr(tyr))
End Function

Function DOPut(S, X, r, q, tyr, sigma, Sb)
    Dim eqt, ert, NDOne, NDTwo, NDThree, NDFour, NDFive, NDSix, NDSeven, NDEight, a, b
    eqt = Exp(-q * tyr)
    ert = Exp(-r * tyr)
    a = (Sb / S) ^ (-1 + (2 * r / sigma ^ 2))
    b = (Sb / S) ^ (1 + (2 * r / sigma ^ 2))
    If S > 0 And X > 0 And tyr > 0 And sigma > 0 Then
        NDOne = Application.NormSDist(BSDOne(S, X, r, q, tyr, sigma))
        NDTwo = Application.NormSDist(BSDTwo(S, X, r, q, tyr, sigma))
        NDThree = Application.NormSDist(BSDThree(S, X, r, q, tyr, sigma, Sb))
        NDFour = Application.NormSDist(BSDFour(S, X, r, q, tyr, sigma, Sb))
        NDFive = Application.NormSDist(BSDFive(S, X, r, q, tyr, sigma, Sb))
        NDSix = Application.NormSDist(BSDSix(S, X, r, q, tyr, sigma, Sb))
        NDSeven = Application.NormSDist(BSDSeven(S, X, r, q, tyr, sigma, Sb))
        NDEight = Application.NormSDist(BSDEight(S, X, r, q, tyr, sigma, Sb))
        DOPut = X * ert * (NDFour - NDTwo - a * (NDSeven - NDFive)) - S * eqt * (NDThree - NDOne - b * (NDEight - NDSix))
    Else
       DOPut = -1
    End If
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

 Function AssetPaths(S, r, q, tyr, sigma, nstep, nsim)
     Dim dt, rnmut, sigt
     Dim i, j As Integer
     Dim spath()
     Randomize
     dt = tyr / nstep
     rnmut = (r - q - 0.5 * sigma ^ 2) * dt
     sigt = sigma * Sqr(dt)
     ReDim spath(nstep, 1 To nsim)
     

     For j = 1 To nsim
        spath(0, j) = S
        For i = 1 To nstep
            randns = Application.NormSInv(Rnd)
            spath(i, j) = spath(i - 1, j) * Exp(rnmut + randns * sigt)
        Next i
     Next j

     AssetPaths = spath
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
 Function AssetPathsHalton(S, r, q, tyr, sigma, nstep, nsim)
     Dim dt, rnmut, sigt
     Dim i, j As Integer
     Dim spath()
     Randomize
     dt = tyr / nstep
     rnmut = (r - q - 0.5 * sigma ^ 2) * dt
     sigt = sigma * Sqr(dt)
     ReDim spath(nstep, 1 To nsim)
     

     For j = 1 To nsim
        spath(0, j) = S
        For i = 1 To nstep
            randns = Application.NormSInv(Halton1(16 + i, 7))
            spath(i, j) = spath(i - 1, j) * Exp(rnmut + randns * sigt)
        Next i
     Next j

     AssetPathsHalton = spath
 End Function
 Sub shareprice_1()
    Dim S, r, q, tyr, sigma, nstep, nsim
    Dim i, j As Integer
    Dim spath()
    Range("A19:z1000").Select
    Selection.ClearContents
    S = Cells(4, 2)
    r = Cells(6, 2)
    q = Cells(8, 2)
    tyr = Cells(11, 2)
    sigma = Cells(12, 2)
    nsim = Cells(14, 2)
    nstep = Cells(15, 2)
    ReDim spath(nstep, 1 To nsim)
    
    Cells(21, 1) = Cells(14, 1)
    Cells(22, 1) = 0
    spath = AssetPaths(S, r, q, tyr, sigma, nstep, nsim)

    For j = 1 To nsim
        Cells(21, 4 + j) = "S" & j
        Cells(22, 4 + j) = S
        For i = 1 To nstep
            Cells(22 + i, 4 + j) = spath(i, j)
            Cells(22 + i, 1) = i
        Next i
    Next j
End Sub
Function DOPutMC_2(S, X, r, q, tyr, sigma, Sb, nstep, nsim)
    Dim payoff, sum, cross
    Dim temp(1)
    Dim spath()
    ReDim spath(nstep, 1 To nsim)
    sum = 0
    cross = 0
    spath = AssetPaths(S, r, q, tyr, sigma, nstep, nsim)
    
    For j = 1 To nsim
        payoff = Application.Max(X - spath(nstep, j), 0)
        For i = 1 To nstep
            If spath(i, j) <= Sb Then
            payoff = 0
            i = nstep
            cross = cross + 1
            End If
        Next i
        sum = sum + payoff
    Next j

    temp(0) = Exp(-r * tyr) * sum / nsim
    temp(1) = cross
    DOPutMC_2 = temp
End Function
Function AsianMC_S(iopt, S, X, r, q, tyr, sigma, nstep, nsim)
    Dim payoff, sum, tot
    Dim spath()
    ReDim spath(nstep, 1 To nsim)
    Dim i, j As Integer

    sum = 0
    spath = AssetPaths(S, r, q, tyr, sigma, nstep, nsim)
    
    For j = 1 To nsim
        tot = 0
        For i = 1 To nstep
            tot = tot + spath(i, j)
        Next i
        payoff = Application.Max(iopt * (tot / nstep - X), 0)
        sum = sum + payoff
    Next j

    AsianMC_S = Exp(-r * tyr) * sum / nsim
End Function


Function AsianMC_K(iopt, S, X, r, q, tyr, sigma, nstep, nsim)
    Dim payoff, sum, tot
    Dim spath()
    ReDim spath(nstep, 1 To nsim)
    Dim i, j As Integer

    sum = 0
    spath = AssetPaths(S, r, q, tyr, sigma, nstep, nsim)
    
    For j = 1 To nsim
        tot = 0
        For i = 1 To nstep
            tot = tot + spath(i, j)
        Next i
        payoff = Application.Max(iopt * (spath(nstep, j) - tot / nstep), 0)
        sum = sum + payoff
    Next j

    AsianMC_K = Exp(-r * tyr) * sum / nsim
End Function

Function LookBackMC_SMax(S, X, r, q, tyr, sigma, nstep, nsim)
    Dim payoff, sum, maxs
    Dim spath()
    ReDim spath(nstep, 1 To nsim)
    Dim i, j As Integer

    sum = 0
    spath = AssetPaths(S, r, q, tyr, sigma, nstep, nsim)
    
    For j = 1 To nsim
        maxs = 0
        For i = 1 To nstep
           If spath(i, j) > maxs Then
           maxs = spath(i, j)
           Else: maxs = maxs
           End If
        Next i
        payoff = Application.Max(maxs - X, 0)
        sum = sum + payoff
    Next j

     LookBackMC_SMax = Exp(-r * tyr) * sum / nsim
End Function

Function LookBackMC_SMin(S, X, r, q, tyr, sigma, nstep, nsim)
    Dim payoff, sum, mins
    Dim spath()
    ReDim spath(nstep, 1 To nsim)
    Dim i, j As Integer

    sum = 0
    spath = AssetPaths(S, r, q, tyr, sigma, nstep, nsim)
    
    For j = 1 To nsim
        mins = spath(1, j)
        For i = 1 To nstep
           If spath(i, j) < mins Then
           mins = spath(i, j)
           Else: mins = mins
           End If
        Next i
        payoff = Application.Max(spath(nstep, j) - mins, 0)
        sum = sum + payoff
    Next j

     LookBackMC_SMin = Exp(-r * tyr) * sum / nsim
End Function



