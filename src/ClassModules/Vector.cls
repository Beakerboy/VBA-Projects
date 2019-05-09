' Class: Vector
'
Dim dA() As Variant 'A column of data dA(intM, 1)
Dim intM As Integer  'number of elements

Private Sub Class_Initialize()
End Sub

Public Property Let Vec(vA As Variant)
    dA = vA
    intM = UBound(vA, 1)
End Property

Public Property Get M() As Integer
    M = intM
End Property

Public Property Get Vec() As Variant
    Vec = dA
End Property

Public Function getValue(i)
    getValue = dA(i, 1)
End Function

Public Function Magnitude()
    Sum = 0
    For i = 1 To Me.M
        Sum = Sum + dA(i, 1) ^ 2
    Next i
    Magnitude = Math.sqr(Sum)
End Function

Public Function DotProduct(oInput As Vector) As Double
    Dim vProduct() As Double
    newM = oInput.M
    If Me.M <> newM Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="DotProduct requires equal length Vectors"
    End If
    vInput = oInput.Vec
    Dim Sum As Double
    Sum = 0
    For i = 1 To Me.M
        Sum = Sum + dA(i, 1) * vInput(i, 1)
    Next i

    DotProduct = Sum
End Function

Public Function Average() As Double
    Average = WorksheetFunction.Average(dA)
End Function

Public Function Stdev() As Double
    Stdev = WorksheetFunction.Stdev(dA)
End Function

Public Function Sum() As Double
    Sum = WorksheetFunction.Sum(dA)
End Function

Public Function ScalarDivide(inNumber) As Vector
    Dim vResult As Variant
    ReDim vResult(1 To M, 1 To 1)
    For i = 1 To M
        vResult(i, 1) = dA(i, 1) / inNumber
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set ScalarDivide = oResult
End Function

Public Function ScalarMultiply(inNumber) As Vector
    Dim vResult As Variant
    ReDim vResult(1 To M, 1 To 1)
    For i = 1 To M
        vResult(i, 1) = dA(i, 1) * inNumber
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set ScalarMultiply = oResult
End Function

Public Function sqr()
    Dim vResult As Variant
    ReDim vResult(1 To M, 1 To 1)
    For i = 1 To M
        vResult(i, 1) = Math.sqr(dA(i, 1))
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set sqr = oResult
End Function

Public Function Inverse()
    Dim vResult As Variant
    ReDim vResult(1 To M, 1 To 1)
    For i = 1 To M
        vResult(i, 1) = 1 / dA(i, 1)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set sqr = oResult
End Function

Public Function toRowMatrix() As Matrix
    Dim oResult As New Matrix
    oResult.Mat = dA
    Set toRowMatrix = oResult.Transpose()
End Function
