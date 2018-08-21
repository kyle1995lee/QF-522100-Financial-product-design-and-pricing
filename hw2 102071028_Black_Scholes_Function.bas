Attribute VB_Name = "Module1"
Function BS(s, k, v, r, t, d)
    d1 = (Log(s / k) + (r - d + v * v / 2) * t) / (v * t ^ 0.5)
    d2 = d1 - v * t ^ 0.5
    BS = s * Exp(-d * t) * Application.NormSDist(d1) - k * Exp(-r * t) * Application.NormSDist(d2)
End Function

