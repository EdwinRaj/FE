Attribute VB_Name = "Milstein"
Function GetStockReturns(ByVal previousReturn As Double, _
                       ByVal interestRate As Double, _
                       ByVal variance As Double, _
                       ByVal randomVariable As Double, _
                       ByVal intervalType As String) As Double

Dim dt As Double
If UCase(intervalType) = "DAILY" Then
dt = 1 / 250
Else
dt = 1 / 2500
End If


'Note: VBA Sqr is a square root function
GetStockReturns = previousReturn + ((interestRate - (variance / 2)) * timeToMaturity) + Sqr(variance * dt) * randomVariable

End Function


Function GetVariance(ByVal previousVariance As Double, _
                       ByVal avgVariance As Double, _
                       ByVal lamda As Double, _
                       ByVal eta As Double, _
                       ByVal randomVariable As Double, _
                       ByVal intervalType As String) As Double

Dim dt As Double
If UCase(intervalType) = "DAILY" Then
dt = 1 / 250
Else
dt = 1 / 2500
End If

Dim lambdaPart As Double
Dim etaPart As Double
Dim etaSquarePart As Double
Dim result As Double

lambdaPart = lamda * (previousVariance - avgVariance) * dt
etaPart = eta * Sqr(previousVariance * dt) * randomVariable
etaSquarePart = (eta * eta) * dt * ((randomVariable * randomVariable) - 1) / 4
 
result = previousVariance - lambdaPart + etaPart + etaSquarePart

'using reflection scheme: absolute value of variance is taken
'If result < 0 Then
'result = 0
'End If

GetVariance = Abs(result)
End Function


