Attribute VB_Name = "Module2"
Sub Black_Scholes()
    s = Range("B2")
    k = Range("B3")
    v = Range("B4")
    r = Range("B5")
    t = Range("B6")
    d = Range("B7")
    If (v > 0) Then
        d1 = (Log(s / k) + (r - d + v * v / 2) * t) / (v * t ^ 0.5)
        d2 = d1 - v * t ^ 0.5
        black = s * Exp(-d * t) * Application.NormSDist(d1) - k * Exp(-r * t) * Application.NormSDist(d2)
        Range("D2") = black
    Else
        MsgBox ("Negative Volatility!")
        Range("D1") = CVErr(xlErrRef)
    End If

End Sub


