Attribute VB_Name = "ViolinChart"
' Function: KernelDensity
' Calculate the kernel density of a set of data at a point, x
'
' Parameters:
'   x         - The point at with the densitity is to be calculated
'   Data      - A vector of data
'   kernel    - Name of the preferred kernel type
'   bandwidth - Name of the bandwidth algorithm to use or the numerical value
'
' Return:
' The kernel density at x.
Public Function KernelDensity(x, Data, Optional kernel = "gaussian", Optional bandwidth = "Silverman")
    n = WorksheetFunction.Count(Data)
    s = WorksheetFunction.StDev_S(Data)

    If bandwidth = "Silverman" Then
        'Silverman's Rule
        bandwidth = s * (4 / 3 / n) ^ 0.2
    ElseIf bandwidth = "Scott" Then
        'Scott's Rule
        bandwidth = s * n ^ (-1 / 5)
    End If
    Sum = 0
    For i = 1 To n
        k = (x - Data(i,1)) / bandwidth
        If kernel = "gaussian" Then
            kernelValue = gaussianKernel(k)
        ElseIf kernel = "uniform" Then
            kernelValue = uniformKernel(k)
        ElseIf kernel = "triangular" Then
            kernelValue = triangularKernel(k)
        ElseIf kernel = "epanechnikov" Then
            kernelValue = epanechnikovKernel(k)
        ElseIf kernel = "quartic" Then
            kernelValue = quarticKernel(k)
        ElseIf kernel = "triweight" Then
            kernelValue = triweightKernel(k)
        ElseIf kernel = "tricube" Then
            kernelValue = tricubeKernel(k)
        End If
  
        Sum = Sum + kernelValue
    Next i
    KernelDensity = Sum / n / bandwidth
End Function

' Function: KernelDensityFromHist
' Produce a Kernel Density given a column of values and a column of frequencies
'
' Parameters:
'   x         - The point at with the densitity is to be calculated
'   Data      - 
'   kernel    - Name of the preferred kernel type
'   bandwidth - Name of the bandwidth algorithm to use or the numerical value
'
' Return:
' The kernel density at x.
Public Function KernelDensityFromHist(x, Data, Optional kernel = "gaussian", Optional bandwidth = "Silverman")
    n = WorksheetFunction.Count(Data) / 2
    Sum = 0
    For i = 1 To n
        Sum = Sum + Data(i, 2)
    Next i
    Dim NewData As Variant
    ReDim NewData(1 To Sum, 1 To 1)
    Dim ArrayCount
    ArrayCount = 1
    For i = 1 To n
        For j = 1 To Data(i, 2)
            NewData(ArrayCount, 1) = Data(i, 1)
            ArrayCount = ArrayCount + 1
        Next j
    Next i
  
    KernelDensityFromHist = KernelDensity(x, NewData, kernel, bandwidth)
End Function

' Function: uniformKernel
Private Function uniformKernel(x)
    If Abs(x) > 1 Then
        k = 0
    Else
        k = 0.5
    End If
    UniformKernel = k
End Function

' Function: TriangularKernel
Private Function TriangularKernel(x)
    If Abs(x) > 1 Then
        k = 0
    Else
        k = 1 - Abs(x)
    End If
    TriangularKernel = k
End Function

' Function: EpanechnikovKernel
Private Function EpanechnikovKernel(x)
    If Abs(x) > 1 Then
        k = 0
    Else
        k = 0.75 * (1 - x ^ 2)
    End If
    EpanechnikovKernel = k
End Function

' Function: QuarticKernel
Private Function QuarticKernel(x)
    If Abs(x) > 1 Then
        k = 0
    Else
        k = 15 / 16 * (1 - x ^ 2) ^ 2
    End If
    QuarticKernel = k
End Function

' Function TriweightKernel
Private Function TriweightKernel(x)
    If Abs(x) > 1 Then
        k = 0
    Else
        k = 35 / 32 * (1 - x ^ 2) ^ 3
    End If
    TriweightKernel = k
End Function

' Function: TricubeKernel
Private Function TricubeKernel(x)
    If Abs(x) > 1 Then
        k = 0
    Else
        k = 70 / 31 * (1 - Abs(x) ^ 3) ^ 3
    End If
    TricubeKernel = k
End Function

' Function: GaussianKernel
Private Function GaussianKernel(x)
    GaussianKernel = WorksheetFunction.Norm_S_Dist(x, False)
End Function

' Function: Violin
' Create a Violin Chart
'
' Parameters:
'   Data          - a range of data
'   XorY          - return the X or Y range
'   Position      - for multiple datasets, where to position the center on the x axis
'   ScalingFactor - The amount to scale the width of the violin to prevent overlap
'
' Returns:
'    An array of Values.
'
' ToDo: Possibly combine X and Y vectors into a 2D list. I can't see why someone would want only one.
Public Function Violin(Data, Optional XorY = "Y", Optional Position = 1, Optional ScalingFactor = 1)
    mu = WorksheetFunction.Average(Data)
    sigma = WorksheetFunction.StDev(Data)
    Dim YVector As Variant
    ReDim YVector(1 To 82, 1 To 1)
    y = mu - 4 * sigma
    For i = 1 To 41
        YVector(i, 1) = y
        YVector(83 - i, 1) = y
        y = y + sigma / 5
    Next i
    If XorY = "Y" Then
        Violin = YVector
    Else
        Dim XVector As Variant
        ReDim XVector(1 To 82, 1 To 1)
        For i = 1 To 41
            x = KernelDensity(YVector(i, 1), Data) / ScalingFactor / 3
            XVector(i, 1) = Position - x
            XVector(83 - i, 1) = Position + x
        Next i
        Violin = XVector
    End If
End Function

' Function: LogViolin
' Create a Violin Chart
'
' Parameters:
'   Data          - a range of data
'   XorY          - return the X or Y range
'   Position      - for multiple datasets, where to position the center on the x axis
'   ScalingFactor - The amount to scale the width of the violin to prevent overlap
'
' Returns:
'    An array of Values.
'
' ToDo: Possibly combine X and Y vectors into a 2D list. I can't see why someone would want only one.
Public Function LogViolin(Data, Optional XorY = "Y", Optional Position = 1, Optional ScalingFactor = 1)
    mu = WorksheetFunction.Average(Data)
    sigma = WorksheetFunction.StDev(Data)
    mu_l = Math.Log(mu / Math.Sqr(1 + sigma ^ 2 / mu ^ 2))
    sigma_l = Math.Sqr(Math.Log(1 + sigma ^ 2 / mu ^ 2))
    ' Take the log of each value
    Dim Data_L As Variant
    Size = WorksheetFunction.Count(Data)
    ReDim Data_L(1 To Size, 1 To 1)
    For i = 1 To Size
        Data_L(i, 1) = Math.Log(Data(i))
    Next i
    Dim YVector As Variant
    ReDim YVector(1 To 72, 1 To 1)
    y = mu_l - 4 * sigma_l
    For i = 1 To 36
        YVector(i, 1) = y
        YVector(73 - i, 1) = y
        y = y + sigma_l / 5
    Next i
    If XorY = "Y" Then
        Dim YVector_L As Variant
        ReDim YVector_L(1 To 72, 1 To 1)
        For i = 1 To 72
            YVector_L(i, 1) = Math.Exp(YVector(i, 1))
        Next i
        LogViolin = YVector_L
    Else
        Dim XVector As Variant
        ReDim XVector(1 To 72, 1 To 1)
        For i = 1 To 36
            x = KernelDensity(YVector(i, 1), Data_L) / ScalingFactor / 3 / Math.Exp(YVector(i, 1))
            XVector(i, 1) = Position - x
            XVector(73 - i, 1) = Position + x
        Next i
        LogViolin = XVector
    End If
End Function
