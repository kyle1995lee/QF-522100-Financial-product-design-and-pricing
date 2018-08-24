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

