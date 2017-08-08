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
  k = (x - Data(i)) / bandwidth
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


Private Function uniformKernel(x)
  If Abs(x) > 1 Then
    k = 0
  Else
    k = 0.5
  End If
  uniformKernel = k
End Function

Private Function triangularKernel(x)
  If Abs(x) > 1 Then
    k = 0
  Else
    k = 1 - Abs(x)
  End If
  triangularKernel = k
End Function


Private Function epanechnikovKernel(x)
  If Abs(x) > 1 Then
    k = 0
  Else
    k = 0.75 * (1 - x ^ 2)
  End If
  epanechnikovKernel = k
End Function

Private Function quarticKernel(x)
  If Abs(x) > 1 Then
    k = 0
  Else
    k = 15 / 16 * (1 - x ^ 2) ^ 2
  End If
  quarticKernel = k
End Function

Private Function triweightKernel(x)
  If Abs(x) > 1 Then
    k = 0
  Else
    k = 35 / 32 * (1 - x ^ 2) ^ 3
  End If
  triweightKernel = k
End Function

Private Function tricubeKernel(x)
  If Abs(x) > 1 Then
    k = 0
  Else
    k = 70 / 31 * (1 - Abs(x) ^ 3) ^ 3
  End If
  tricubeKernel = k
End Function

Private Function gaussianKernel(x)
  gaussianKernel = WorksheetFunction.Norm_S_Dist(x, False)
End Function

'Create a Violin Chart
Public Function Violin(Data, Optional XorY = "Y", Optional LeftorRight = "Left", Optional Position = 1, Optional ScalingFactor = 1)
' Data: a range of data
' XorY: return the X or Y range
' LeftorRight: return the data to the left or right of the position
' Position: for multiple datasets, where to position the center
' ScalingFactor: The amount to scale the width of the violin to prevent overlap
mu = WorksheetFunction.Average(Data)
sigma = WorksheetFunction.StDev(Data)
Dim YVector As Variant
ReDim YVector(1 To 41, 1 To 1)
y = mu - 4 * sigma
For i = 1 To 41
  YVector(i, 1) = y
  y = y + sigma / 5
Next i
If XorY = "Y" Then
  Violin = YVector
Else
  Dim XVector As Variant
  ReDim XVector(1 To 41, 1 To 1)
  For i = 1 To 41
    x = KernelDensity(YVector(i, 1), Data) / ScalingFactor / 3
    If LeftorRight = "Left" Then
      XVector(i, 1) = Position - x
    Else
      XVector(i, 1) = Position + x
    End If
  Next i
  Violin = XVector
End If

End Function
