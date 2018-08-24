Attribute VB_Name = "Trinomial"
Sub TrinomialTree()
    Range("A20:Z38").Select
    Selection.ClearContents
    Range("B21").Select
    Dim delt, rnmut, u, d
    Dim i As Integer, j As Integer
    
    S = Cells(4, 4)
    K = Cells(5, 4)
    tyr = Cells(11, 4)
    nstep = Cells(15, 4)
    lamda = Cells(18, 4)
    delt = tyr / nstep
    r = Cells(6, 4)
    q = Cells(8, 4)
    sigma = Cells(13, 4)
    rnmut = (r - q - 0.5 * sigma ^ 2) * delt
    ermqdt = Exp((r - q) * delt)
    u = Exp(lamda * sigma * Sqr(delt))
    d = 1 / u
    Cells(21, 2) = S
    

    
    xm = 21
    ym = 2
    
    For i = 1 To 2 * nstep
        Cells(i + xm, 0 + ym) = ""
    Next i
    For j = 1 To nstep
        For i = 0 To 2 * j
            Cells(i + xm, j + ym) = u ^ j * d ^ i * S
        Next i
    For i = 2 * j + 1 To 2 * nstep
            Cells(i + xm, j + ym) = ""
        Next i
    Next j
    
    For i = 0 To nstep
        Cells(20, 2 + i) = i
    Next i
    For i = 0 To 2 * nstep
        Cells(21 + i, 1) = i
    Next i

End Sub

Function TriOptionValue(iopt, iea, S, X, r, q, tyr, sigma, nstep, lamda)
'   Returns Binomial Option Value (imod=1 for CRR, 2 for LR; iea=1 for euro, 2 for amer)
'   Uses BSDOne fn
'   Uses BSDTwo fn
'   Uses PPNormInv fn
    Dim delt, erdt, ermqdt, u, d, d1, d2, p, pdash, pstar
    Dim i As Integer, j As Integer
    Dim vvec() As Variant
    If imod = 2 Then nstep = Application.Odd(nstep)
    
    ReDim vvec(2 * nstep)
    
    If S > 0 And X > 0 And tyr > 0 And sigma > 0 Then
    delt = tyr / nstep
    erdt = Exp(r * delt)
    ermqdt = Exp((r - q) * delt)
'   Choice between imod=0 (JR), imod=1 (Cox,Ross&Rubinstein) and imod=2 (Leisen&Reimer)
    If imod = 0 Then
      rnmut = (r - q - 0.5 * sigma ^ 2) * delt
      
u = Exp(lamda * sigma * Sqr(delt))
d = 1 / u
pu = 1 / (2 * lamda ^ 2) + (r - sigma ^ 2 / 2) * Sqr(delt) / (2 * lamda * sigma)
pm = 1 - 1 / (lamda ^ 2)
pd = 1 - pu - pm
    For i = 0 To 2 * nstep
      vvec(i) = Application.Max(iopt * (S * (d ^ nstep) * (u ^ i) - X), 0)
Next i
For j = nstep - 1 To 0 Step -1
            For i = 0 To 2 * j
      vvec(i) = (pu * vvec(i + 2) + pm * vvec(i + 1) + pd * vvec(i)) / erdt
      If iea = 2 Then vvec(i) = Application.Max(vvec(i), iopt * (S * (u ^ i) * (d ^ j) - X))
            Next i
Next j
    TriOptionValue = vvec(0)
    Else
    TriOptionValue = -1
    End If
    End If
    
End Function

