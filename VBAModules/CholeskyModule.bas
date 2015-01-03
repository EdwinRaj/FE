Attribute VB_Name = "CholeskyModule"
Function Cholesky(a)

b = a.Value
n = UBound(b, 1)

ReDim L(1 To n, 1 To n)

L(1, 1) = 1

For i = 2 To n
 Sum2 = 0
 For j = 1 To i - 1
  Sum1 = 0
  For k = 1 To j - 1
   Sum1 = Sum1 + L(i, k) * L(j, k)
  Next k
  L(i, j) = (1 / L(j, j)) * (b(i, j) - Sum1)
 Next j
  For k = 1 To i - 1
   Sum2 = Sum2 + L(i, k) * L(i, k)
  Next k
 L(i, i) = Sqr(b(i, i) - Sum2)
Next i

Cholesky = L

End Function


