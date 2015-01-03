Attribute VB_Name = "HestonModel"
'Heston Call Price by Monte Carlo Simulation
'Sourced from the book by ROUAH & VAINBERG
'Modified to factor daily and hourly factors
Function HestonMC(kappa, theta, lambda, rho, SigmaV, daynum, startS, r, startv, k, ITER)
Dim allS() As Double, Stock() As Double

simPath = 0
ReDim allS(daynum) As Double, Stock(ITER) As Double
deltat = (1 / 365)

For itcount = 1 To ITER
    lnSt = Log(startS)
    lnvt = Log(startv)
    curv = startv
    curS = startS
        For daycnt = 1 To daynum
            e = Application.NormSInv(Rnd)
            eS = Application.NormSInv(Rnd)
            ev = rho * eS + Sqr(1 - rho ^ 2) * e
            'update the stock price
            lnSt = lnSt + (r - 0.5 * curv) * deltat + Sqr(curv) * Sqr(deltat) * eS
            curS = Exp(lnSt)
            lnvt = lnvt + (kappa * (theta - curv) - lambda * curv - 0.5 * SigmaV) * deltat + SigmaV * (1 / Sqr(curv)) * Sqr(deltat) * ev
            curv = Exp(lnvt)
            allS(daycnt) = curS
        Next daycnt
    simPath = simPath + Exp((-daynum / 365) * r) * Application.Max(allS(daynum) - k, 0)
Next itcount
  HestonMC = simPath / ITER
End Function


