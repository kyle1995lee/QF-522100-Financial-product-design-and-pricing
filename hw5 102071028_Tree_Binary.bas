Attribute VB_Name = "Module3"
Sub Binomial_Tree_Userform()
    UserForm1.Show
End Sub

Sub GoHome()
    Sheets("Home").Select
End Sub


Sub Go_Simplified()
    Sheets("Simplified").Select
End Sub



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

Function BinOptionValue(imod, iopt, iea, S, X, r, q, tyr, sigma, nstep)
'   Returns Binomial Option Value (imod=1 for CRR, 2 for LR; iea=1 for euro, 2 for amer)
'   Uses BSDOne fn
'   Uses BSDTwo fn
'   Uses PPNormInv fn
    Dim delt, erdt, ermqdt, u, d, d1, d2, p, pdash, pstar
    Dim i As Integer, j As Integer
    Dim vvec() As Variant
    If imod = 2 Then nstep = Application.Odd(nstep)
    ReDim vvec(nstep)
    If S > 0 And X > 0 And tyr > 0 And sigma > 0 Then
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
      d2 = BSDTwo(S, X, r, q, tyr, sigma)
      d1 = BSDOne(S, X, r, q, tyr, sigma)
      p = PPNormInv(d2, nstep)
      pstar = 1 - p
      pdash = PPNormInv(d1, nstep)
      u = ermqdt * pdash / p
      d = (ermqdt - p * u) / (1 - p)
      End If
    End If
    

    
    For i = 0 To nstep
        If Application.Max(iopt * (S * (u ^ i) * (d ^ (nstep - i))), 0) < 25 Then
            vvec(i) = 1000
        ElseIf Application.Max(iopt * (S * (u ^ i) * (d ^ (nstep - i))), 0) >= 40 Then
            vvec(i) = 1000 + (40 - 25) * 170
        Else
            vvec(i) = 1000 + (Application.Max(iopt * (S * (u ^ i) * (d ^ (nstep - i)) - 25), 0)) * 170
        End If
    Next i
    
    For j = nstep - 1 To 0 Step -1
        For i = 0 To j
            vvec(i) = (p * vvec(i + 1) + pstar * vvec(i)) / erdt
            If iea = 2 Then vvec(i) = Application.Max(vvec(i), iopt * (S * (u ^ i) * (d ^ (j - i)) - X))
        Next i
    Next j
    
    BinOptionValue = vvec(0)
    Else
    BinOptionValue = -1
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


Sub Simplified_Tree()
n = Range("n")
u = Range("u")
p0 = Range("p0")

Range("A8:ZA30").Select
Selection.ClearContents

Dim t As Integer
For t = 2 To 2 + n
    Cells(8, t) = -2 + t
    Cells(t + 7, 1) = -2 + t
Next t


Range("B9").Select
Dim i As Integer
Dim j As Integer
Cells(9, 2) = p0

For i = 2 To n + 1
    Cells(9, i + 1) = Cells(9, i) + u
    For j = 2 To i
        Cells(j + 8, i + 1) = Cells(j + 7, i) - u
    Next j
Next i

End Sub

