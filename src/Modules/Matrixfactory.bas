Public Function IdentityMatrix(M As Integer) As Matrix
    Dim vA() As Variant
    ReDim vA(1 To M, 1 To M)
    For i = 1 To M
        vA(i, i) = 1
    Next
    Dim oA As Matrix
    
    Set oA = New Matrix
    oA.Mat = vA
    Set IdentityMatrix = oA
End Function

Public Function ZeroMatrix(M As Integer, N As Integer, Optional init = 0) As Matrix
    Dim vA() As Double
    ReDim vA(1 To M, 1 To M)
    If init <> 0 Then
        For i = 1 To M
            For j = 1 To N
                vA(i, j) = init
            Next j
        Next i
    End If
    Dim oA As Matrix
    Set oA = New Matrix
    oA.Mat = vA
    Set Identity = oA
End Function

Public Function IdentityMatrixExcel(M As Integer)
    Dim oA As Matrix
    Set oA = IdentityMatrix(M)
    vA = oA.Mat
    IdentityMatrixExcel = vA
End Function
