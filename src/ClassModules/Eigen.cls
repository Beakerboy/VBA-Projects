Dim oMatrix As Matrix
Dim oEigenVec As Matrix
Dim oEigenVal As Vector

Public Property Set Matrix(inMatrix)
    Set oMatrix = inMatrix
End Property

Public Property Get EigenVec() As Matrix
    Set EigenVec = oEigenVec
End Property

Public Property Get EigenVal() As Vector
    Set EigenVal = oEigenVal
End Property

Public Sub PolynomialRoots()
    Covariance = oMatrix.Mat
    A = 1
    b = -1 * Covariance(1, 1) - Covariance(2, 2)
    c = -1 * Covariance(2, 1) * Covariance(1, 2) + Covariance(1, 1) * Covariance(2, 2)

    Eigenvalue1 = (-b - Math.sqr(b * b - 4 * A * c)) / 2 / A
    Eigenvalue2 = (-1 * b + Math.sqr(b * b - 4 * A * c)) / 2 / A

    Dim Eigen As Variant
    ReDim Eigen(1 To 2, 1 To 1)
    Eigen(1, 1) = Eigenvalue1
    Eigen(2, 1) = Eigenvalue2
    Set oEigenVal = New Vector
    oEigenVal.Vec = Eigen
    
    Dim vMatrix As Variant
    ReDim vMatrix(1 To 2, 1 To 2)

    Dim oLambda As Matrix
    Set oLambda = Identity(2).ScalarMultiply(oEigenVal.getValue(2))
    
    Dim oAMinusLambda As Matrix
    Set oAMinusLambda = oMatrix.Subtract(oLambda)
    A21 = oAMinusLambda.getValue(1, 2)
    A11 = oAMinusLambda.getValue(1, 1)
    vMatrix(2, 1) = Math.sqr(A21 ^ 2 / (A11 ^ 2 + A21 ^ 2))
    vMatrix(1, 2) = vMatrix(2, 1)
    vMatrix(1, 1) = vMatrix(2, 1) * A11 / A21
    vMatrix(2, 2) = -1 * vMatrix(1, 1)
    
    Set oEigenVec = New Matrix
    oEigenVec.Mat = vMatrix
End Sub

Function JK() As Variant
    Dim iVariables As Integer
    iVariables = oMatrix.M
    ReDim vEigenVal(1 To iVariables, 1 To 1)
    
    Dim Ematrix As Variant
    Dim i As Long, j As Long, k As Long, iter As Long, p As Long
    Dim den As Double, hold As Double, num As Double
    Dim Test As Double, tmp As Double
    Const eps As Double = 1E-16
    
    On Error GoTo EndProc
    Dim Orig_A As Matrix
    Dim A As Matrix
    Set Orig_A = oMatrix
    Set A = oMatrix
    p = A.M
    ReDim Ematrix(1 To p, 1 To p)
    Dim x As Vector
    Dim y As Vector
    For iter = 1 To 500
        'Orthogonalize pairs of columns in upper off diag
        For i = 1 To p - 1
            For j = i + 1 To p
                Set x = A.GetColumn(i)
                Set y = A.GetColumn(j)
                num = 2 * x.DotProduct(y)
                den = x.DotProduct(x) - y.DotProduct(y)
                
                'Skip rotation if aij is zero and correct ordering
                If Abs(num) < eps And den >= 0 Then Exit For
                Dim oRotationMatrix As Matrix
                Set oRotationMatrix = JKRotationMatrix(num, den)
                Dim oxMat As New Matrix
                Dim oyMat As New Matrix
                oxMat.Mat = x.Vec
                oyMat.Mat = y.Vec
                Set oTwoColumns = oxMat.AugmentRight(oyMat)
                Set oRotatedTwoColumns = oTwoColumns.Multiply(oRotationMatrix)
                Set A = A.ReplaceColumn(i, oRotatedTwoColumns.GetColumn(1))
                Set A = A.ReplaceColumn(j, oRotatedTwoColumns.GetColumn(2))
            Next j
        Next i
        
        'Test for convergence
        Test = A.Multiply(A.Transpose).Trace
        If Abs(Test - hold) < eps And iter > 5 Then Exit For
        hold = Test
    Next iter
    
    If iter = 101 Then MsgBox "JK Iteration has not converged."
    Dim Eval As Matrix
    Set Eval = A.Transpose().Multiply(Orig_A).Multiply(A)
    Set oEigenVec = A
    'Compute eigenvalues/eigenvectors
    'Normalize eigenvectors
            ' oSigns = Eval.getDiagonalElements
            ' Set oEigenVal = New Vector
            ' Set oEigenVal = Eval.getDiagonalElements.HadamardProduct(oSigns).Pow(1/3).HadamardProduct(oSigns)
            ' Set oSigns = A.getRow(1).sgn()
            ' oScalingFactors = oEigenVal.abs().HadamardProduct(oSigns)
            ' Set oEigenVec = oEigenVec.ColumnDivide(oScalingFactors)
    Dim oEigenVector As Vector
    Dim oScaledVector As Vector
    For i = 1 To p
        'Compute eigenvalues
        iSign = Math.Sgn(Eval.getValue(i, i))
        vEigenVal(i, 1) = iSign * (iSign * Eval.getValue(i, i)) ^ (1 / 3)
            
        'iSign = Math.Sgn(A.getValue(1, i))
        iSign = 1
        Set oEigenVector = A.GetColumn(i)
        Set oScaledVector = oEigenVector.ScalarDivide(iSign * Abs(vEigenVal(i, 1)))
        Set oEigenVec = oEigenVec.ReplaceColumn(i, oScaledVector)

    Next i
    Set oEigenVal = New Vector
    oEigenVal.Vec = vEigenVal
    Exit Function
    
EndProc:
        MsgBox prompt:="Error in function PCA.Eigen()!" & vbCr & vbCr & _
            "Error: " & Err.Description & ".", Buttons:=48, _
            Title:="Run time error!"
End Function

Private Function JKRotationMatrix(num, den) As Matrix
    Dim oRotationMatrix As New Matrix
    'Perform Rotation
    If Abs(num) <= Abs(den) Then
        Tan2 = Abs(num) / Abs(den)          ': eq. 11
        Cos2 = 1 / sqr(1 + Tan2 * Tan2)     ': eq. 12
        Sin2 = Tan2 * Cos2                  ': eq. 13
    Else
        Cot2 = Abs(den) / Abs(num)          ': eq. 16
        Sin2 = 1 / sqr(1 + Cot2 * Cot2)     ': eq. 17
        Cos2 = Cot2 * Sin2                  ': eq. 18
    End If
                
    Cos_ = sqr((1 + Cos2) / 2)              ': eq. 14/19
    Sin_ = Sin2 / (2 * Cos_)                ': eq. 15/20
                
    If den < 0 Then
        tmp = Cos_
        Cos_ = Sin_                         ': table 21
        Sin_ = tmp
    End If
                
    Sin_ = Math.Sgn(num) * Sin_                  ': sign table 21
               
    'Rotate
    Dim vRotationMatrix As Variant
    ReDim vRotationMatrix(1 To 2, 1 To 2)
    vRotationMatrix(1, 1) = Cos_
    vRotationMatrix(1, 2) = -1 * Sin_
    vRotationMatrix(2, 1) = Sin_
    vRotationMatrix(2, 2) = Cos_
    oRotationMatrix.Mat = vRotationMatrix
    Set JKRotationMatrix = oRotationMatrix
End Function
