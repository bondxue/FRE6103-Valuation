Attribute VB_Name = "Module1"
Option Explicit
'PV of annuity
Function PVA(C As Double, n As Integer, r As Double, F As Double, Optional s As Integer = 0) As Double
'C: amount of single cash flow
'n: date of final payment
'r: discount rate
'F: face value
's: starting date of first payment, taking time 0 as the present

    If r = 0 Then
        PVA = n * C + F
    Else
        PVA_C = C / r * (1 - 1 / (1 + r) ^ n)
        PVA_F = F / (1 + r) ^ n
        PVA = (PVA_C + PVA_F) / (1 + r) ^ (s - 1)
    End If
    
End Function

'PV of growing annuity
Function PVGA(C1 As Double, n As Integer, r As Double, g As Double, F As Double, Optional s As Integer = 0) As Double
'C1: the starting payment
'n: date of final payment
'r: discount rate
'g: growth rate in cash flow
'F: face value
's: the time period of first cash flow

    If r = 0 Then
        PVGA = C1 / g * ((1 + g) ^ n - 1) + F
    Else
        r_star = (r - g) / (1 + g)
        C0 = C1 / (1 + g)
        PVGA = PVA(C0, n, r_star, F, 1) / ((1 + r / 100) ^ (s - 1))
    End If
    
End Function

'PV of linearly growing perpetuity
Function PVLGP(a As Double, b As Double, r As Double, s As Integer) As Double
'a: initial payment
'b increase in payment from one period to the next as constant. a+bt
'r: discount rate
's: the first payment date

    PVLGP = ((a + b * s) / r + b / (r) ^ 2) / (1 + r) ^ (s - 1)

End Function
'PV of linearly growing annuity
Function PVLGA(a As Double, b As Double, n As Integer, r As Double, s As Integer) As Double
'a: initial payment
'b increase in payment from one period to the next as constant. a+bt
'n: number of payments
'r: discount rate
's: the first payment date

    PVLGA = (PVLGP(a, b, r, s) - PVLGP(a, b, r, n + s))

End Function
'PV of mean-reverting annuity
Function PVMRA(C1 As Double, kappa As Double, mu As Double, n As Double, r As Double, s As Integer) As Double
'C1: first payment
'mu: long-run mean cash flow
'kappa: periodic mean-reverting speed
'n: number of payments
'r: discount rate
's: the first payment date

    PVRMA = PVA(mu, n, r, 0, s) + PVGA(C1 - mu, n, r, -kappa, 0, s)

End Function
