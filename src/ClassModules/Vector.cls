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
