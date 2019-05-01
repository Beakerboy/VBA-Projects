Dim dA() As Variant
Dim intM As Integer 'rows
Dim intN As Integer 'columns

Private Sub Class_Initialize()
End Sub

Public Property Let Mat(vA As Variant)
    dA = vA
    intM = UBound(vA, 1)
    intN = UBound(vA, 2)
End Property

Public Property Get N() As Integer
    N = intN
End Property

Public Property Get M() As Integer
    M = intM
End Property

Public Property Get Mat() As Variant
    Mat = dA
End Property
Public Function Multiply(oInput As Matrix) As Matrix
    If Me.N <> oInput.M Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Multiplication requires compatible sized matricies"
    End If
    Dim vProduct() As Variant
    
    ReDim vProduct(1 To Me.M, 1 To oInput.N)
    Dim i As Integer
    For i = 1 To Me.M
        Dim oRow As Vector
        Set oRow = Me.GetRow(i)
        Dim j As Integer
        For j = 1 To oInput.N
            Dim oColumn As Vector
            Set oColumn = oInput.GetColumn(j)
            vProduct(i, j) = oRow.DotProduct(oColumn)
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vProduct
    Set Multiply = oResult
End Function

Public Function ScalarDivide(dInput As Double) As Matrix
    Dim i As Integer
    Dim vProduct() As Variant
    
    ReDim vProduct(1 To Me.M, 1 To Me.N)
    For i = 1 To Me.M
        Dim j As Integer
        For j = 1 To Me.N
            vProduct(i, j) = dA(i, j) / dInput
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vProduct
    Set ScalarDivide = oResult
End Function

Public Function Transpose() As Matrix
    Dim oResult As New Matrix
    oResult.Mat = WorksheetFunction.Transpose(dA)
    Set Transpose = oResult
End Function

Public Function GetRow(M As Integer) As Vector
    Dim vResult() As Variant
    ReDim vResult(1 To intN, 1 To 1)
    For i = 1 To intN
        vResult(i, 1) = dA(M, i)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set GetRow = oResult
End Function

Public Function GetColumn(N As Integer) As Vector
    Dim vResult() As Variant
    ReDim vResult(1 To intM, 1 To 1)
    For i = 1 To intM
        vResult(i, 1) = dA(i, N)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set GetColumn = oResult
End Function

Public Function getValue(i, j)
    getValue = dA(i, j)
End Function
