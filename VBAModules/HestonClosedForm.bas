Attribute VB_Name = "HestonClosedForm"
Option Base 1

Private Const fSteps = 400
Private Const fLength = 80

' ---------------------------------- COMPLEX LIBRARY ------------------------

Public Type cNum
    rP As Double
    iP As Double
End Type
Function thePI()
    thePI = Application.Pi()
End Function
Function set_cNum(rPart, iPart) As cNum
    set_cNum.rP = rPart
    set_cNum.iP = iPart
End Function
Function cNumProd(cNum1 As cNum, cNum2 As cNum) As cNum
    cNumProd.rP = (cNum1.rP * cNum2.rP) - (cNum1.iP * cNum2.iP)
    cNumProd.iP = (cNum1.rP * cNum2.iP) + (cNum1.iP * cNum2.rP)
End Function
Function cNumConj(cNum1 As cNum) As cNum
    cNumConj.rP = cNum1.rP
    cNumConj.iP = -cNum1.iP
End Function
Function cNumDiv(cNum1 As cNum, cNum2 As cNum) As cNum
    Dim conj_cNum2 As cNum
    conj_cNum2 = cNumConj(cNum2)
    cNumDiv.rP = (cNum1.rP * conj_cNum2.rP - cNum1.iP * conj_cNum2.iP) / (cNum2.rP ^ 2 + cNum2.iP ^ 2)
    cNumDiv.iP = (cNum1.rP * conj_cNum2.iP + cNum1.iP * conj_cNum2.rP) / (cNum2.rP ^ 2 + cNum2.iP ^ 2)
End Function
Function cNumAdd(cNum1 As cNum, cNum2 As cNum) As cNum
    cNumAdd.rP = cNum1.rP + cNum2.rP
    cNumAdd.iP = cNum1.iP + cNum2.iP
End Function
Function cNumSub(cNum1 As cNum, cNum2 As cNum) As cNum
    cNumSub.rP = cNum1.rP - cNum2.rP
    cNumSub.iP = cNum1.iP - cNum2.iP
End Function
Function cNumSqrt(cNum1 As cNum) As cNum
    r = Sqr(cNum1.rP ^ 2 + cNum1.iP ^ 2)
    y = Atn(cNum1.iP / cNum1.rP)
    cNumSqrt.rP = Sqr(r) * Cos(y / 2)
    cNumSqrt.iP = Sqr(r) * Sin(y / 2)
End Function
Function cNumPower(cNum1 As cNum, n As Double) As cNum
    r = Sqr(cNum1.rP ^ 2 + cNum1.iP ^ 2)
    y = Atn(cNum1.iP / cNum1.rP)
    cNumPower.rP = r ^ n * Cos(y * n)
    cNumPower.iP = r ^ n * Sin(y * n)
End Function
Function cNumExp(cNum1 As cNum) As cNum
    cNumExp.rP = Exp(cNum1.rP) * Cos(cNum1.iP)
    cNumExp.iP = Exp(cNum1.rP) * Sin(cNum1.iP)
End Function
Function cNumSq(cNum1 As cNum) As cNum
    cNumSq = cNumProd(cNum1, cNum1)
End Function
Function cNumReal(cNum1 As cNum) As Double
    cNumReal = cNum1.rP
End Function
Function cNumLn(cNum1 As cNum) As cNum
    r = (cNum1.rP ^ 2 + cNum1.iP ^ 2) ^ 0.5
    theta = Atn(cNum1.iP / cNum1.rP)
    cNumLn.rP = Application.Ln(r)
    cNumLn.iP = theta
End Function
Function cNumPowercNum(cNum1 As cNum, cNum2 As cNum) As cNum
    r = Sqr(cNum1.rP ^ 2 + cNum1.iP ^ 2)
    y = Atn(cNum1.iP / cNum1.rP)
    cNumPowercNum.rP = r ^ cNum2.rP * Exp(-cNum2.iP * y) * Cos(cNum2.rP * y + cNum2.iP * Log(r))
    cNumPowercNum.iP = r ^ cNum2.rP * Exp(-cNum2.iP * y) * Sin(cNum2.rP * y + cNum2.iP * Log(r))
End Function

' Lewis (2000) Integrand

Function intH(k As cNum, X, V0, tau, thet, kappa, SigmaV, rho, gam) As Double
Dim b As cNum, im As cNum, thetaadj As cNum, c As cNum
Dim d As cNum, f As cNum, h As cNum, AA As cNum, BB As cNum
Dim Hval As cNum, t As cNum, a As cNum, re As cNum

' Lewis Parameters

omega = kappa * thet
ksi = SigmaV
theta = kappa

t = set_cNum(ksi ^ 2 * tau / 2, 0)
a = set_cNum(2 * omega / ksi ^ 2, 0)

If (gam = 1) Then
    thetaadj = set_cNum(theta, 0)
Else
    thetaadj = set_cNum((1 - gam) * rho * ksi + Sqr(theta ^ 2 - gam * (1 - gam) * ksi ^ 2), 0)
End If

im = set_cNum(0, 1)
re = set_cNum(1, 0)

b = cNumProd(set_cNum(2, 0), cNumDiv(cNumAdd(thetaadj, cNumProd(im, cNumProd(k, set_cNum(rho * ksi, 0)))), set_cNum(ksi ^ 2, 0)))
c = cNumDiv(cNumSub(cNumSq(k), cNumProd(im, k)), set_cNum(ksi ^ 2, 0))
d = cNumSqrt(cNumAdd(cNumSq(b), cNumProd(set_cNum(4, 0), c)))
f = cNumDiv(cNumAdd(b, d), set_cNum(2, 0))
h = cNumDiv(cNumAdd(b, d), cNumSub(b, d))
AA = cNumSub(cNumProd(cNumProd(f, a), t), cNumProd(a, cNumLn(cNumDiv(cNumSub(re, cNumProd(h, cNumExp(cNumProd(d, t)))), cNumSub(re, h)))))
BB = cNumDiv(cNumProd(f, cNumSub(re, cNumExp(cNumProd(d, t)))), cNumSub(re, cNumProd(h, cNumExp(cNumProd(d, t)))))
Hval = cNumExp(cNumAdd(AA, cNumProd(BB, set_cNum(V0, 0))))
intH = cNumReal(cNumProd(cNumDiv(cNumExp(cNumProd(cNumProd(set_cNum(-X, 0), im), k)), cNumSub(cNumSq(k), cNumProd(im, k))), Hval))
End Function

' Heston Price by Fundamental Transform

Function HCTrans(S, k, r, delta, V0, tau, ki, thet, kappa, SigmaV, rho, gam, PutCall As String)
Dim int_x() As Double, int_y() As Double
Dim pass_phi As cNum

' Lewis Parameters

omega = kappa * thet
ksi = SigmaV
theta = kappa
kmax = Round(Application.Max(1000, 10 / Sqr(V0 * tau)), 0)
ReDim int_x(kmax * 5) As Double, int_y(kmax * 5) As Double

X = Application.Ln(S / k) + (r - delta) * tau

cnt = 0
For phi = 0.000001 To kmax Step 0.2

cnt = cnt + 1
    int_x(cnt) = phi
    pass_phi = set_cNum(phi, ki)
    int_y(cnt) = intH(pass_phi, X, V0, tau, thet, kappa, SigmaV, rho, gam)
Next phi

CallPrice = (S * Exp(-delta * tau) - (1 / thePI) * k * Exp(-r * tau) * TRAPnumint(int_x, int_y))

If PutCall = "Call" Then
    HCTrans = CallPrice
ElseIf PutCall = "Put" Then
    HCTrans = CallPrice + k * Exp(-r * tau) - S * Exp(-delta * tau)
End If

End Function

' Trapezoidal Rule

Function TRAPnumint(X, y) As Double
Dim n As Integer, t As Integer
    n = Application.Count(X)
    TRAPnumint = 0
    For t = 2 To n
        TRAPnumint = TRAPnumint + 0.5 * (X(t) - X(t - 1)) * (y(t - 1) + y(t))
    Next
End Function



