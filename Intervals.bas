Public Function ConfInt(x, Ys, Xs, alpha, Optional RTO = False)
'x: evaluation point
'Ys: array of y values
'Xs: array of x values
'alpha: interval size, for 95% confidence, alpha=.05
'RTO: if true, use regression through origin

Count = WorksheetFunction.Count(Xs)
'v=degrees of freedom
v = Count
If RTO Then
  v = Count - 1
Else
  v = Count - 2
End If
SSx = WorksheetFunction.SumSq(Xs)
DevSq = WorksheetFunction.DevSq(Xs)
SqDev = (x - WorksheetFunction.Average(Xs)) ^ 2
t = WorksheetFunction.T_Inv_2T(alpha, v)

If RTO Then
  Slope = WorksheetFunction.SumProduct(Xs, Ys) / SSx
  SStot = WorksheetFunction.SumSq(Ys)
  Yhat = WorksheetFunction.MMult(Xs, Slope)
  SSreg = WorksheetFunction.SumSq(Yhat)
  SSres = SStot - SSreg
  StEyx = (SSres / v) ^ 0.5
  ConfInt = x * t * StEyx / SSx ^ 0.5
Else
  StEyx = WorksheetFunction.StEyx(Ys, Xs)
  ConfInt = t * StEyx * (1 / Count + SqDev / DevSq) ^ 0.5
End If
End Function

Function PredInt(x, Ys, Xs, alpha, Optional RTO = False, Optional q = 1)
'x: evaluation point
'Ys: array of y values
'Xs: array of x values
'alpha: interval size, for 95% confidence, alpha=.05
'RTO: if true, use regression through origin
'q: number of replicates

Count = WorksheetFunction.Count(Xs)
'v=degrees of freedom
v = Count
If RTO Then
  v = Count - 1
Else
  v = Count - 2
End If
SSx = WorksheetFunction.SumSq(Xs)
DevSq = WorksheetFunction.DevSq(Xs)
SqDev = (x - WorksheetFunction.Average(Xs)) ^ 2
t = WorksheetFunction.T_Inv_2T(alpha, v)
StEyx = WorksheetFunction.StEyx(Ys, Xs)
If RTO Then
  Slope = WorksheetFunction.SumProduct(Xs, Ys) / SSx
  SStot = WorksheetFunction.SumSq(Ys)
  Yhat = WorksheetFunction.MMult(Xs, Slope)
  SSreg = WorksheetFunction.SumSq(Yhat)
  SSres = SStot - SSreg
  StEyx = (SSres / v) ^ 0.5
  PredInt = t * StEyx * (1 / q + x ^ 2 / SSx) ^ 0.5
Else
  StEyx = WorksheetFunction.StEyx(Ys, Xs)
  PredInt = t * StEyx * (1 / q + 1 / Count + SqDev / DevSq) ^ 0.5
End If
End Function
Function FiducialInt(Yo, Ys, Xs, alpha, constant, UpperLower)
'constant:true=calc normally; false=force to zero
'This is the inverse of the confidence interval. Given a Y value, what is the range on the x.
n = WorksheetFunction.Count(Xs)
v = n - 2
If constant = False Then v = v + 1
t = WorksheetFunction.T_Inv_2T(alpha, v)
LinEst = WorksheetFunction.LinEst(Ys, Xs, constant, False)
b1 = WorksheetFunction.Index(LinEst, 1)

StEyx = WorksheetFunction.StEyx(Ys, Xs)
DevSq = WorksheetFunction.DevSq(Xs)

Xbar = WorksheetFunction.Average(Xs)
Ybar = WorksheetFunction.Average(Ys)
Xo = (Yo - b0) / b1
Xu = Xbar + (b1 * (Yo - Ybar) + t * StEyx * ((Yo - Ybar) ^ 2 / DevSq + b1 ^ 2 / n - t ^ 2 * StEyx ^ 2 / n / DevSq) ^ 0.5) / (b1 ^ 2 - t ^ 2 * StEyx ^ 2 / DevSq)
Xl = Xbar + (b1 * (Yo - Ybar) - t * StEyx * ((Yo - Ybar) ^ 2 / DevSq + b1 ^ 2 / n - t ^ 2 * StEyx ^ 2 / n / DevSq) ^ 0.5) / (b1 ^ 2 - t ^ 2 * StEyx ^ 2 / DevSq)
If UpperLower = 1 Then FiducialInt = Xu Else FiducialInt = Xl

End Function
Function InverseConfInt(Yo, Ys, Xs, alpha, constant, UpperLower)
'constant:true=calc normally; false=force to zero
'This is the inverse of the confidence interval. Given a Y value, what is the range on the x.
n = WorksheetFunction.Count(Xs)
v = n - 2
If constant = False Then v = v + 1
t = WorksheetFunction.T_Inv_2T(alpha, v)
LinEst = WorksheetFunction.LinEst(Ys, Xs, constant, False)
b1 = WorksheetFunction.Index(LinEst, 1)
Sum = WorksheetFunction.Index(LinEst, 3, 2) ^ 2 / WorksheetFunction.Index(LinEst, 2, 1) ^ 2

S = WorksheetFunction.Index(LinEst, 3, 2)


Xbar = WorksheetFunction.Average(Xs)
Ybar = WorksheetFunction.Average(Ys)
Xo = (Yo - b0) / b1
Part1 = b1 * (Yo - Ybar)
Part2 = t * S * ((Yo - Ybar) ^ 2 / Sum + b1 ^ 2 / n - t ^ 2 * S ^ 2 / n / Sum) ^ 0.5
Part3 = b1 ^ 2 - t ^ 2 * S ^ 2 / Sum

Part4 = t * S / Sum ^ 0.5
If constant Then Xu = Xbar + (Part1 + Part2) / Part3 Else Xu = Yo / (b + Part4)
If constant Then Xl = Xbar + (Part1 - Part2) / Part3 Else Xl = Yo / (b - Part4)
If UpperLower = 1 Then InverseConfInt = Xu Else InverseConfInt = Xl

End Function

Function InversePredInt(Yo, Ys, Xs, alpha, constant, UpperLower, Optional q = 1)
'constant:true=calc normally; false=force to zero
'This is the inverse of the prediction interval. Given a Y value, what is the range on the x.
n = WorksheetFunction.Count(Xs)
v = n - 2
If constant = False Then v = v + 1
t = WorksheetFunction.T_Inv_2T(alpha, v)
LinEst = WorksheetFunction.LinEst(Ys, Xs, constant, True)
b1 = WorksheetFunction.Index(LinEst, 1, 1)
beta = WorksheetFunction.Index(LinEst, 2, 1)
S = WorksheetFunction.Index(LinEst, 3, 2)

Xbar = WorksheetFunction.Average(Xs)
Ybar = WorksheetFunction.Average(Ys)
Part1 = b1 * (Yo - Ybar)
Part2 = t * S * ((Yo - Ybar) ^ 2 * beta ^ 2 / S ^ 2 + b1 ^ 2 * (n + q) / n / q - t ^ 2 * beta ^ 2 * (n + q) / n / q) ^ 0.5
Part3 = b1 ^ 2 - t ^ 2 * beta ^ 2

Part4 = t * S * (Yo ^ 2 * beta ^ 2 / S ^ 2 + Part3 / q) ^ 0.5

If constant Then Xu = Xbar + (Part1 + Part2) / Part3 Else Xu = (Yo * b1 - Part4) / Part3
If constant Then Xl = Xbar + (Part1 - Part2) / Part3 Else Xl = (Yo * b1 + Part4) / Part3
If UpperLower = 1 Then InversePredInt = Xu Else InversePredInt = Xl

End Function
Public Function HatMatrix(Xs, index1, index2)
Xt = WorksheetFunction.Transpose(Xs)
XtX = WorksheetFunction.MMult(Xt, Xs)

H = WorksheetFunction.MMult(Xs, WorksheetFunction.MMult(WorksheetFunction.MInverse(XtX), Xt))
HatMatrix = H(index1, index2)
End Function
