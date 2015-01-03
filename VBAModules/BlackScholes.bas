Attribute VB_Name = "BlackScholes"
Function BSprice(pc, S, k, vol, d, r, t)

'pc  put/call indicator call=1, put=-1
'S   Stock price at 0
'K   strike
'vol volatility
'd   dividend yield
'r   riskless rate
't   time to maturity


d1 = (Log(S / k) + t * (r - d + (vol ^ 2) / 2)) / (vol * Sqr(t))
d2 = d1 - vol * Sqr(t)

BSprice = pc * Exp(-d * t) * S * Application.WorksheetFunction.NormSDist(pc * d1) - pc * k * Exp(-r * t) * Application.WorksheetFunction.NormSDist(pc * d2)

End Function
Function BSdelta(pc, S, k, vol, d, r, t)

'pc  put/call indicator call=1, put=-1
'S   Stock price at 0
'K   strike
'vol volatility
'd   dividend yield
'r   riskless rate
't   time to maturity


d1 = (Log(S / k) + t * (r - d + (vol ^ 2) / 2)) / (vol * Sqr(t))

If pc = 1 Then
BSdelta = Exp(-d * t) * Application.WorksheetFunction.NormSDist(d1)
Else
BSdelta = Exp(-d * t) * (Application.WorksheetFunction.NormSDist(d1) - 1)
End If

End Function

Function BSgamma(pc, S, k, vol, d, r, t)

'pc  put/call indicator call=1, put=-1
'S   Stock price at 0
'K   strike
'vol volatility
'd   dividend yield
'r   riskless rate
't   time to maturity

d1 = (Log(S / k) + t * (r - d + (vol ^ 2) / 2)) / (vol * Sqr(t))

BSgamma = Exp(-d * t) * Exp((-d1 ^ 2) / 2) / (Sqr(2 * Application.WorksheetFunction.Pi()) * S * vol * Sqr(t))

End Function

Function BStheta(pc, S, k, vol, d, r, t)

'pc  put/call indicator call=1, put=-1
'S   Stock price at 0
'K   strike
'vol volatility
'd   dividend yield
'r   riskless rate
't   time to maturity

d1 = (Log(S / k) + t * (r - d + (vol ^ 2) / 2)) / (vol * Sqr(t))
d2 = d1 - vol * Sqr(t)

BStheta = -Exp(-d * t) * Exp((-d1 ^ 2) / 2) * S * vol / (Sqr(2 * Application.WorksheetFunction.Pi()) * 2 * Sqr(t)) + pc * d * S * Exp(-q * t) * Application.WorksheetFunction.NormSDist(pc * d1) - pc * r * k * Exp(-r * t) * Application.WorksheetFunction.NormSDist(pc * d2)

End Function

Function BSvega(pc, S, k, vol, d, r, t)

'pc  put/call indicator call=1, put=-1
'S   Stock price at 0
'K   strike
'vol volatility
'd   dividend yield
'r   riskless rate
't   time to maturity

d1 = (Log(S / k) + t * (r - d + (vol ^ 2) / 2)) / (vol * Sqr(t))

BSvega = Exp(-d * t) * S * Sqr(t) * Exp((-d1 ^ 2) / 2) / (Sqr(2 * Application.WorksheetFunction.Pi()))

End Function

Function BSrho(pc, S, k, vol, d, r, t)

'pc  put/call indicator call=1, put=-1
'S   Stock price at 0
'K   strike
'vol volatility
'd   dividend yield
'r   riskless rate
't   time to maturity

d1 = (Log(S / k) + t * (r - d + (vol ^ 2) / 2)) / (vol * Sqr(t))
d2 = d1 - vol * Sqr(t)

BSrho = pc * k * t * Exp(-r * t) * Application.WorksheetFunction.NormSDist(pc * d2)

End Function

Function BSvol(pc, S, k, price, d, r, t, Optional start)

'pc    put/call indicator call=1, put=-1
'S     Stock price at 0
'K     strike
'price option premium
'd     dividend yield
'r     riskless rate
't     time to maturity
'start starting value for vol, optional, by default=0.2

If IsMissing(start) Then
  start = 0.2
End If

voli = start
pricei = BSprice(pc, S, k, voli, d, r, t)
vegai = BSvega(pc, S, k, voli, d, r, t)

Do Until Abs(price - pricei) < 0.000001

voli = voli + (price - pricei) / vegai
pricei = BSprice(pc, S, k, voli, d, r, t)
vegai = BSvega(pc, S, k, voli, d, r, t)

Loop

BSvol = voli

End Function






