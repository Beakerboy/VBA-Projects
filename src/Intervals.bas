' Statistical Intervals
'
' ConfInt        - The Confidence Interval of y=mx+b or y=mx
' PredInt        - The Prediction Interval of y=mx+b or y=mx
' QuadConfInt    - The Confidence Interval of y-mx²+bx
' QuadPredInt    - The Prediction Interval of y-mx²+bx
' InverseConfInt - The Inverse of the Confidence Interval
' InversePredInt - The Inverse of the Prediction Interval

' Function: ConfInt
' Calculate the confidence interval at a point given data
'
' Parameters:
'   x     - evaluation point
'   Ys    - array of y values
'   Xs    - array of x values
'   alpha - interval size, for 95% confidence, alpha=.05
'   SLR   - if true, use simple linear regression
'         - if false use regression through the origin
'
' Returns:
'   An interval in the Y direction above or below the line of best-fit
Public Function ConfInt(x, Ys, Xs, alpha, Optional SLR = TRUE)
    Count = WorksheetFunction.Count(Xs)
    'v=degrees of freedom
    v = Count
    If SLR Then
        v = Count - 2
    Else
        v = Count - 1
    End If
    SSx = WorksheetFunction.SumSq(Xs)
    DevSq = WorksheetFunction.DevSq(Xs)
    SqDev = (x - WorksheetFunction.Average(Xs)) ^ 2
    t = WorksheetFunction.T_Inv_2T(alpha, v)
    If SLR Then
        StEyx = WorksheetFunction.StEyx(Ys, Xs)
        ConfInt = t * StEyx * (1 / Count + SqDev / DevSq) ^ 0.5
    Else
        Slope = WorksheetFunction.SumProduct(Xs, Ys) / SSx
        SStot = WorksheetFunction.SumSq(Ys)
        Yhat = WorksheetFunction.MMult(Xs, Slope)
        SSreg = WorksheetFunction.SumSq(Yhat)
        SSres = SStot - SSreg
        StEyx = (SSres / v) ^ 0.5
        ConfInt = x * t * StEyx / SSx ^ 0.5
    End If
End Function

' Function: PredInt
' Calculate the prediction interval at a point given data
'
' Parameters:
'   x     - evaluation point
'   Ys    - array of y values
'   Xs    - array of x values
'   alpha - interval size, for 95% confidence, alpha=.05
'   SLR   - if true, use regression through origin
'         - if false use regression through the origin
'   q     - number of replicates
'
' Returns:
'   An interval in the Y direction above or below the regression line of best-fit
Function PredInt(x, Ys, Xs, alpha, Optional SLR = TRUE, Optional q = 1)
    Count = WorksheetFunction.Count(Xs)
    'v=degrees of freedom
    v = Count
    If SLR Then
        v = Count - 2
    Else
        v = Count - 1
    End If
    SSx = WorksheetFunction.SumSq(Xs)
    DevSq = WorksheetFunction.DevSq(Xs)
    SqDev = (x - WorksheetFunction.Average(Xs)) ^ 2
    t = WorksheetFunction.T_Inv_2T(alpha, v)
    StEyx = WorksheetFunction.StEyx(Ys, Xs)
    If SLR Then
        StEyx = WorksheetFunction.StEyx(Ys, Xs)
        PredInt = t * StEyx * (1 / q + 1 / Count + SqDev / DevSq) ^ 0.5
    Else
        Slope = WorksheetFunction.SumProduct(Xs, Ys) / SSx
        SStot = WorksheetFunction.SumSq(Ys)
        Yhat = WorksheetFunction.MMult(Xs, Slope)
        SSreg = WorksheetFunction.SumSq(Yhat)
        SSres = SStot - SSreg
        StEyx = (SSres / v) ^ 0.5
        PredInt = t * StEyx * (1 / q + x ^ 2 / SSx) ^ 0.5
    End If
End Function

' Function QuadConfInt
' Calculate the confidence interval on a 2nd order polynomial with zero intercept
'
' Parameters:
'   x     - evaluation point
'   Ys    - array of y values
'   Xs    - array of x values
'   alpha - interval size, for 95% confidence, alpha=.05
'
' Returns:
'   An interval in the Y direction above or below the line of best-fit
Function QuadConfInt(X, Ys, Xs, alpha)
    count = WorksheetFunction.count(Xs)
    v = count - 2
    SumX2 = WorksheetFunction.SumSq(Xs)
    SumX3 = WorksheetFunction.SumProduct(Xs, Xs, Xs)
    SumX4 = WorksheetFunction.SumProduct(Xs, Xs, Xs, Xs)
    SumYX = WorksheetFunction.SumProduct(Ys, Xs)
    SumYX2 = WorksheetFunction.SumProduct(Ys, Xs, Xs)
    det = SumX4 * SumX2 - SumX3 ^ 2
    m = (SumX2 * SumYX2 - SumX3 * SumYX) / det
   
    b = (SumX4 * SumYX - SumX3 * SumYX2) / det
    SStot = WorksheetFunction.SumSq(Ys)
    SSreg = SumX4 * m ^ 2 + 2 * SumX3 * m * b + SumX2 * b * b
    SSres = SStot - SSreg
    StEyx = (SSres / v) ^ 0.5
    se = (X ^ 2 * SumX2 - 2 * X * SumX3 + SumX4) / det
    T = WorksheetFunction.T_Inv_2T(alpha, v)
    QuadConfInt = T * X * StEyx * se ^ 0.5
End Function

' Function QuadPredInt
' Calculate the prediction interval on a 2nd order polynomial with zero intercept
'
' Parameters:
'   x     - evaluation point
'   Ys    - array of y values
'   Xs    - array of x values
'   alpha - interval size, for 95% confidence, alpha=.05
'
' Returns:
'   An interval in the Y direction above or below the line of best-fit
Function QuadPredInt(X, Ys, Xs, alpha)
    count = WorksheetFunction.count(Xs)
    v = count - 2
    SumX2 = WorksheetFunction.SumSq(Xs)
    SumX3 = WorksheetFunction.SumProduct(Xs, Xs, Xs)
    SumX4 = WorksheetFunction.SumProduct(Xs, Xs, Xs, Xs)
    SumYX = WorksheetFunction.SumProduct(Ys, Xs)
    SumYX2 = WorksheetFunction.SumProduct(Ys, Xs, Xs)
    det = SumX4 * SumX2 - SumX3 ^ 2
    m = (SumX2 * SumYX2 - SumX3 * SumYX) / det
   
    b = (SumX4 * SumYX - SumX3 * SumYX2) / det
    SStot = WorksheetFunction.SumSq(Ys)
    SSreg = SumX4 * m ^ 2 + 2 * SumX3 * m * b + SumX2 * b * b
    SSres = SStot - SSreg
    StEyx = (SSres / v) ^ 0.5
    se = (X ^ 2 * SumX2 - 2 * X * SumX3 + SumX4) / det
    T = WorksheetFunction.T_Inv_2T(alpha, v)
    QuadPredInt = T * StEyx * (X ^ 2 * se + 1) ^ 0.5
End Function

' Function: FiducialInt
' This is the inverse of the confidence interval. Given a Y value, what is the range on the x.
'
' Parameters:
'   Yo         - 
'   Ys         - 
'   Xs         - 
'   alpha      - 
'   constant   - true=calc normally; false=force to zero
'   UpperLower - 1 = upper interval, -1 = lower interval
'
' Returns:
'   The interval in the x direction above or below the SLR line of best fit
Function FiducialInt(Yo, Ys, Xs, alpha, constant, UpperLower)
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
    
	DeltaX = Xbar + (b1 * (Yo - Ybar) + UpperLower * t * StEyx * ((Yo - Ybar) ^ 2 / DevSq + b1 ^ 2 / n - t ^ 2 * StEyx ^ 2 / n / DevSq) ^ 0.5) / (b1 ^ 2 - t ^ 2 * StEyx ^ 2 / DevSq)
	FiducialInt = DeltaX
End Function

' Function: InverseConfInt
' Given data and a confidence level, find the x value along either the upper or lower confidence band at a given y value
'
' Prarmeters
'   Yo         - 
'   Ys         - 
'   Xs    - 
'   alpha - 
'   SLR   - if true, use regression through origin
'         - if false use regression through the origin
'   Upper - Return the upper band if TRUE
'         - Return the lower band if FALSE
'
' Returns
' The value on either the upper or lower confidence band there y crosses the band
Function InverseConfInt(Yo, Ys, Xs, alpha, SLR, UpperLower)
    n = WorksheetFunction.Count(Xs)
    v = n - 2
    If SLR = False Then v = v + 1
    t = WorksheetFunction.T_Inv_2T(alpha, v)
    LinEst = WorksheetFunction.LinEst(Ys, Xs, SLR, False)
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
    If SLR Then Xu = Xbar + (Part1 + Part2) / Part3 Else Xu = Yo / (b + Part4)
    If SLR Then Xl = Xbar + (Part1 - Part2) / Part3 Else Xl = Yo / (b - Part4)
	If UpperLower Then InverseConfInt = Xu Else InverseConfInt = Xl
End Function

' Function: InversePredInt
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
                                
' Function: ConfVector
' Return an array of confidence Intervals
Public Function ConfVector(Ys, Xs, alpha, count, plusorminus, Optional RTO = False)
    Xinput = EqualSpace(Xs, count)
    Dim ReturnVector As Variant
    ReDim ReturnVector(1 To count, 1 To 1)
    For i = 1 To count
        If plusorminus = "plus" Then
            ReturnVector(i, 1) = Forecast(Xinput(i, 1), Ys, Xs, RTO) + ConfInt(Xinput(i, 1), Ys, Xs, alpha, RTO)
        Else
            ReturnVector(i, 1) = Forecast(Xinput(i, 1), Ys, Xs, RTO) - ConfInt(Xinput(i, 1), Ys, Xs, alpha, RTO)
        End If
    Next i
    ConfVector = ReturnVector
End Function

' Function: PredVector
' Return an array of prediction Intervals
Public Function PredVector(Ys, Xs, alpha, count, plusorminus, Optional RTO = False)
    Xinput = EqualSpace(Xs, count)
    Dim ReturnVector As Variant
    ReDim ReturnVector(1 To count, 1 To 1)
    For i = 1 To count
        If plusorminus = "plus" Then
            ReturnVector(i, 1) = Forecast(Xinput(i, 1), Ys, Xs, RTO) + PredInt(Xinput(i, 1), Ys, Xs, alpha, RTO)
        Else
            ReturnVector(i, 1) = Forecast(Xinput(i, 1), Ys, Xs, RTO) - PredInt(Xinput(i, 1), Ys, Xs, alpha, RTO)
        End If
    Next i
    PredVector = ReturnVector
End Function

'Return an array of confidence Intervals
Public Function QuadConfVector(Ys, Xs As Range, alpha, count, plusorminus)
Xinput = EqualSpace(Xs, count)
Dim ReturnVector As Variant
ReDim ReturnVector(1 To count, 1 To 1)
For i = 1 To count
    Forecast = QuadForecastVBA(Xinput(i, 1), Ys, Xs)
    Conf = QuadConfInt(Xinput(i, 1), Ys, Xs, alpha)
  If plusorminus = "plus" Then
    ReturnVector(i, 1) = Forecast + Conf
  Else
    ReturnVector(i, 1) = Forecast - Conf
  End If
Next i
QuadConfVector = ReturnVector
End Function

'Return an array of prediction Intervals
Public Function QuadPredVector(Ys, Xs As Range, alpha, count, plusorminus)
Xinput = EqualSpace(Xs, count)
Dim ReturnVector As Variant
ReDim ReturnVector(1 To count, 1 To 1)
For i = 1 To count
    Forecast = QuadForecastVBA(Xinput(i, 1), Ys, Xs)
    Conf = QuadPredInt(Xinput(i, 1), Ys, Xs, alpha)
  If plusorminus = "plus" Then
    ReturnVector(i, 1) = Forecast + Conf
  Else
    ReturnVector(i, 1) = Forecast - Conf
  End If
Next i
QuadPredVector = ReturnVector
End Function

' Function: HatMatrix
Public Function HatMatrix(Xs, index1, index2)
    Xt = WorksheetFunction.Transpose(Xs)
    XtX = WorksheetFunction.MMult(Xt, Xs)

    H = WorksheetFunction.MMult(Xs, WorksheetFunction.MMult(WorksheetFunction.MInverse(XtX), Xt))
    HatMatrix = H(index1, index2)
End Function

' Function: Leverage
Public Function Leverage(Xs, index)
    Leverage = HatMatrix(Xs, index, index)
End Function

' Function: EqualSpace
' Return an array of {count} number of points eqaly spaced along the span of VectorObject
Public Function EqualSpace(VectorObject, count)
    Max = WorksheetFunction.Max(VectorObject)
    Min = WorksheetFunction.Min(VectorObject)

    Delta = (Max - Min) / (count - 1)
    Dim ReturnVector As Variant
    ReDim ReturnVector(1 To count, 1 To 1)
    For i = 1 To count
        ReturnVector(i, 1) = Min + Delta * (i - 1)
    Next i
    EqualSpace = ReturnVector
End Function

' Function: Forcast
' Wrapper for the Excel Forcast function, but also accepts RTO 
Public Function Forecast(X, Ys, Xs, Optional RTO = False)
    If RTO Then
        LinEst = WorksheetFunction.LinEst(Ys, Xs, False, True)
        Forecast = X * LinEst(1, 1)
    Else
        Forecast = WorksheetFunction.Forecast(X, Ys, Xs)
    End If
End Function

'Function: QuadForcastVBA                       
'Given some data in ranges of X and Y, create a least squares Y=mx²+bx model and apply it to a range of new data, 
'
' Parameters:
'   X     - evaluation points
'   Ys    - Data for the regression
'   Xs    - Data for the regression
'
' Returns:
'   An array of new Y values
Public Function QuadForecastVBA(X, Ys, Xs As Range)
     LinEst = WorksheetFunction.LinEst(Ys, Vander(Xs), False, True)
    QuadForecastVBA = X * LinEst(1, 1) + X ^ 2 * LinEst(1, 2)
End Function                                

