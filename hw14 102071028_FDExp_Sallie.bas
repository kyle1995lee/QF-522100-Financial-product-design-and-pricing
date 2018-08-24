Attribute VB_Name = "Module4"
Function matvalEuPut(S, x, r, T, sigma, Smax, dS, dt)
    Dim M, N
    Dim matval()
    Dim i, j As Integer
    M = Smax / dS
    N = T / dt
    ReDim matval(M, N)
    For i = 0 To M
        matval(i, N) = Application.Max(x - dS * i, 0)
    Next i
    For j = 0 To N
        matval(0, j) = x * Exp(-r * dt * (N - j))
        matval(M, j) = 0
    Next j
    matvalEuPut = matval
End Function
Function ai(r, sigma, dt, M)
    Dim i As Integer
    Dim a()
    ReDim a(M)
    For i = 0 To M
        a(i) = 0.5 * dt * ((sigma * i) ^ 2 - r * i)
    Next i
    ai = a
End Function
Function bi(r, sigma, dt, M)
    Dim i As Integer
    Dim b()
    ReDim b(M)
    For i = 0 To M
        b(i) = 1 - dt * ((sigma * i) ^ 2 + r)
    Next i
    bi = b
End Function
Function ci(r, sigma, dt, M)
    Dim i As Integer
    Dim c()
    ReDim c(M)
    For i = 0 To M
        c(i) = 0.5 * dt * ((sigma * i) ^ 2 + r * i)
    Next i
    ci = c
End Function
Function EuPutExpl(S, x, r, T, sigma, Smax, dS, dt)
    Dim M, N, Mo
    Dim i, j As Integer
    Dim matval(), a(), b(), c()
    M = Int(Smax / dS)
    N = Int(T / dt)
    ReDim matval(M, N), a(M), b(M), c(M)
    matval = matvalEuPut(S, x, r, T, sigma, Smax, dS, dt)
    a = ai(r, sigma, dt, M)
    b = bi(r, sigma, dt, M)
    c = ci(r, sigma, dt, M)
   For j = N - 1 To 0 Step -1
        For i = 1 To M - 1
            matval(i, j) = a(i) * matval(i - 1, j + 1) + b(i) * matval(i, j + 1) + c(i) * matval(i + 1, j + 1)
        Next i
    Next j
    Mo = S / dS
    EuPutExpl = matval(Mo, 0)
End Function
Function EuPutExplall(S, x, r, T, sigma, Smax, dS, dt)
    Dim M, N, Mo
    Dim i, j As Integer
    Dim matval(), a(), b(), c()
    M = Int(Smax / dS)
    N = Int(T / dt)
    ReDim matval(M, N), a(M), b(M), c(M)
    matval = matvalEuPut(S, x, r, T, sigma, Smax, dS, dt)
    a = ai(r, sigma, dt, M)
    b = bi(r, sigma, dt, M)
    c = ci(r, sigma, dt, M)
    
For j = N - 1 To 0 Step -1
        For i = 1 To M - 1
            matval(i, j) = a(i) * matval(i - 1, j + 1) + b(i) * matval(i, j + 1) + c(i) * matval(i + 1, j + 1)
        Next i
    Next j
    EuPutExplall = matval
End Function

Sub EuPutExplThreeD()
    Dim M, N
    Dim i, j As Integer
    Dim matval()
    Range("a18:dz200").Select
    Selection.ClearContents
    Dim S, x, T, r, tyr
    S = Cells(4, 2)
    x = Cells(5, 2)
    r = Cells(6, 2)
    tyr = Cells(11, 2)
    sigma = Cells(12, 2)
    Smax = Cells(13, 2)
    dS = Cells(14, 2)
    dt = Cells(15, 2)
M = Int(Smax / dS)
    N = Int(tyr / dt)
    ReDim matval(M, N)
    matval = EuPutExplall(S, x, r, tyr, sigma, Smax, dS, dt)
    For j = 0 To N
        For i = 0 To M
            Cells(18 + i, 1 + j) = matval(i, j)
        Next i
     Next j
     Mo = S / dS
     Cells(11, 5) = matval(Mo, 0)
End Sub

Function SLMBinOptionValue(imod, iopt, iea, S, x, r, q, tyr, sigma, nstep)
'   Returns Binomial Option Value (imod=1 for CRR, 2 for LR; iea=1 for euro, 2 for amer)
'   Uses BSDOne fn
'   Uses BSDTwo fn
'   Uses PPNormInv fn
    Dim delt, erdt, ermqdt, u, d, d1, d2, p, pdash, pstar
    Dim i As Integer, j As Integer
    Dim vvec() As Variant
    If imod = 2 Then nstep = Application.Odd(nstep)
    ReDim vvec(nstep)
    If S > 0 And x > 0 And tyr > 0 And sigma > 0 Then
    delt = tyr / nstep
    erdt = Exp(r * delt)
    ermqdt = Exp((r - q) * delt)
'   Choice between imod=0 (JR), imod=1 (Cox,Ross&Rubinstein) and imod=2 (Leisen&Reimer)
    If imod = 0 Then
      rnmut = (r - q - 0.5 * sigma ^ 2) * delt
      u = Exp(rnmut + sigma * Sqr(delt))
      d = Exp(rnmut - sigma * Sqr(delt))
      p = 0.5
      pstar = 1 - p
    Else
      If imod = 1 Then
      u = Exp(sigma * Sqr(delt))
      d = 1 / u
      p = (ermqdt - d) / (u - d)
      pstar = 1 - p
      Else
      d2 = BSDTwo(S, x, r, q, tyr, sigma)
      d1 = BSDOne(S, x, r, q, tyr, sigma)
      p = PPNormInv(d2, nstep)
      pstar = 1 - p
      pdash = PPNormInv(d1, nstep)
      u = ermqdt * pdash / p
      d = (ermqdt - p * u) / (1 - p)
      End If
    End If
    For i = 0 To nstep
      vvec(i) = Application.Max(iopt * (50 * (S * (u ^ i) * (d ^ (nstep - i)) - x)) / S, 9.25)
    Next i
        For j = nstep - 1 To 0 Step -1
            For i = 0 To j
            x = 131.75
            If j > 8 Then x = 129.5
            If j > 12 Then x = 127
            If j > 16 Then x = 124.5
      vvec(i) = (p * vvec(i + 1) + pstar * vvec(i)) / erdt
      If iea = 2 Then vvec(i) = Application.Max(vvec(i), iopt * (50 * (S * (u ^ i) * (d ^ (j - i)) - x) / S))
            Next i
        Next j
    SLMBinOptionValue = vvec(0)
    Else
    SLMBinOptionValue = -1
    End If
End Function

Function SLM2BinOptionValue(imod, iopt, iea, S, x, r, q, tyr, sigma, nstep)
'   Returns Binomial Option Value (imod=1 for CRR, 2 for LR; iea=1 for euro, 2 for amer)
'   Uses BSDOne fn
'   Uses BSDTwo fn
'   Uses PPNormInv fn
    Dim delt, erdt, ermqdt, u, d, d1, d2, p, pdash, pstar
    Dim i As Integer, j As Integer
    Dim vvec() As Variant
    If imod = 2 Then nstep = Application.Odd(nstep)
    ReDim vvec(nstep)
    If S > 0 And x > 0 And tyr > 0 And sigma > 0 Then
    delt = tyr / nstep
    erdt = Exp(r * delt)
    ermqdt = Exp((r - q) * delt)
'   Choice between imod=0 (JR), imod=1 (Cox,Ross&Rubinstein) and imod=2 (Leisen&Reimer)
    If imod = 0 Then
      rnmut = (r - q - 0.5 * sigma ^ 2) * delt
      u = Exp(rnmut + sigma * Sqr(delt))
      d = Exp(rnmut - sigma * Sqr(delt))
      p = 0.5
      pstar = 1 - p
    Else
      If imod = 1 Then
      u = Exp(sigma * Sqr(delt))
      d = 1 / u
      p = (ermqdt - d) / (u - d)
      pstar = 1 - p
      Else
      d2 = BSDTwo(S, x, r, q, tyr, sigma)
      d1 = BSDOne(S, x, r, q, tyr, sigma)
      p = PPNormInv(d2, nstep)
      pstar = 1 - p
      pdash = PPNormInv(d1, nstep)
      u = ermqdt * pdash / p
      d = (ermqdt - p * u) / (1 - p)
      End If
    End If
    For i = 0 To nstep
      vvec(i) = Application.Max(iopt * (50 * (1 / S * (u ^ i) * (d ^ (nstep - i)) - 1 / x)) * S, 9.25)
    Next i
        For j = nstep - 1 To 0 Step -1
            For i = 0 To j
            x = 1 / 131.75
            If j > 8 Then x = 1 / 129.5
            If j > 12 Then x = 1 / 127
            If j > 16 Then x = 1 / 124.5
      vvec(i) = (p * vvec(i + 1) + pstar * vvec(i)) / erdt
      If iea = 2 Then vvec(i) = Application.Max(vvec(i), iopt * (50 * (1 / S * (u ^ i) * (d ^ (j - i)) - 1 / x) * S))
            Next i
        Next j
    SLM2BinOptionValue = vvec(0)
    Else
    SLM2BinOptionValue = -1
    End If
End Function

Sub us_show()
    UserForm1.Show
End Sub

Function BinOptionValue(imod, iopt, iea, S, x, r, q, tyr, sigma, nstep)
'   Returns Binomial Option Value (imod=1 for CRR, 2 for LR; iea=1 for euro, 2 for amer)
'   Uses BSDOne fn
'   Uses BSDTwo fn
'   Uses PPNormInv fn
    Dim delt, erdt, ermqdt, u, d, d1, d2, p, pdash, pstar
    Dim i As Integer, j As Integer
    Dim vvec() As Variant
    If imod = 2 Then nstep = Application.Odd(nstep)
    ReDim vvec(nstep)
    If S > 0 And x > 0 And tyr > 0 And sigma > 0 Then
    delt = tyr / nstep
    erdt = Exp(r * delt)
    ermqdt = Exp((r - q) * delt)
'   Choice between imod=0 (JR), imod=1 (Cox,Ross&Rubinstein) and imod=2 (Leisen&Reimer)
    If imod = 0 Then
      rnmut = (r - q - 0.5 * sigma ^ 2) * delt
      u = Exp(rnmut + sigma * Sqr(delt))
      d = Exp(rnmut - sigma * Sqr(delt))
      p = 0.5
      pstar = 1 - p
    Else
      If imod = 1 Then
      u = Exp(sigma * Sqr(delt))
      d = 1 / u
      p = (ermqdt - d) / (u - d)
      pstar = 1 - p
      Else
      d2 = BSDTwo(S, x, r, q, tyr, sigma)
      d1 = BSDOne(S, x, r, q, tyr, sigma)
      p = PPNormInv(d2, nstep)
      pstar = 1 - p
      pdash = PPNormInv(d1, nstep)
      u = ermqdt * pdash / p
      d = (ermqdt - p * u) / (1 - p)
      End If
    End If
    For i = 0 To nstep
      vvec(i) = Application.Max(iopt * (S * (u ^ i) * (d ^ (nstep - i)) - x), 0)
    Next i
        For j = nstep - 1 To 0 Step -1
            For i = 0 To j
      vvec(i) = (p * vvec(i + 1) + pstar * vvec(i)) / erdt
      If iea = 2 Then vvec(i) = Application.Max(vvec(i), iopt * (S * (u ^ i) * (d ^ (j - i)) - x))
            Next i
        Next j
    BinOptionValue = vvec(0)
    Else
    BinOptionValue = -1
    End If
End Function

Function PPNormInv(z, N)
'   Returns the Peizer-Pratt Inversion
'   Only defined for n odd
'   Used in LR Binomial Option Valuation
    Dim c1
    N = Application.Odd(N)
    c1 = Exp(-((z / (N + 1 / 3 + 0.1 / (N + 1))) ^ 2) * (N + 1 / 6))
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





