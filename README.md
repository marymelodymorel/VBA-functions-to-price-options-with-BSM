#put function

Option Explicit

Public Function BSM_put(s As Double, k As Double, r As Double, t As Double, vol As Double) As Variant

Dim d1 As Double
Dim d2 As Double
Dim N_d1 As Double
Dim N_d2 As Double

d1 = (Application.WorksheetFunction.Ln(s / k) + (r + vol ^ (2) / 2) * t) / (vol * Sqr(t))
d2 = d1 - vol * Sqr(t)

N_d1 = Application.WorksheetFunction.Norm_Dist(d1, 0, 1, True)
N_d2 = Application.WorksheetFunction.Norm_Dist(d2, 0, 1, True)

BSM_put = k * Exp(-r * t) * (1 - N_d2) - s * (1 - N_d1)

End Function