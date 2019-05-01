Dim vData As Matrix
Dim vAverages As Variant
Dim vStdevs As Variant
Dim oNormalData As Matrix
Dim oCorrelation As Matrix
Dim oEigenVec As Matrix
Dim vEigenVal As Variant
Dim oTransformed As Matrix

Property Let Data(inData As Variant)
    Set vData = New Matrix
    vData.Mat = inData
End Property

Property Get Averages()
    Averages = vAverages
End Property

Property Get Stdevs()
    Stdevs = vStdevs
End Property

Property Get NormalData()
    NormalData = oNormalData.Mat
End Property

Property Get Correlation()
    Correlation = oCorrelation.Mat
End Property

Property Get EigenVec()
    EigenVec = oEigenVec.Mat
End Property

Property Get EigenVal()
    EigenVal = vEigenVal
End Property

Property Get TransformedData()
    TransformedData = oTransformed.Mat
End Property

Private Sub CalcAves()
    iVariables = vData.N
    ReDim vAverages(1 To iVariables)
    
    Dim i As Integer
    
    For i = 1 To iVariables
        vAverages(i) = vData.GetColumn(i).Average
    Next i

End Sub

Private Sub CalcStdevs()
    iVariables = vData.N
    ReDim vStdevs(1 To iVariables)
    
    Dim i As Integer
    
    For i = 1 To iVariables
        vStdevs(i) = vData.GetColumn(i).Stdev
    Next i
End Sub

Public Sub Normalize()
    CalcAves
    CalcStdevs
    iSamples = vData.M
    iVariables = vData.N
    ReDim vNormalData(1 To iSamples, 1 To iVariables)
    
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To iSamples
        For j = 1 To iVariables
            vNormalData(i, j) = (vData.getValue(i, j) - vAverages(j)) / vStdevs(j)
        Next j
    Next i
    Set oNormalData = New Matrix
    oNormalData.Mat = vNormalData
End Sub

Public Sub CorrelationMatrix()
    iSamples = vData.M
    Dim oCorrelationt As Matrix
    Set oCorrelationt = oNormalData.Transpose()
    Set oCorrelation = oCorrelationt.Multiply(oNormalData).ScalarDivide(iSamples - 1)
End Sub

Function Eigen() As Variant
    M = oCorrelation.Mat
    Dim iVariables As Integer
    iVariables = vData.N
    ReDim vEigenVal(1 To iVariables)
    
    Dim A() As Variant, Ematrix As Variant
    Dim i As Long, j As Long, k As Long, iter As Long, p As Long
    Dim den As Double, hold As Double, Sin_ As Double, num As Double
    Dim Sin2 As Double, Cos2 As Double, Cos_ As Double, Test As Double
    Dim Tan2 As Double, Cot2 As Double, tmp As Double
    Const eps As Double = 1E-16
    
    On Error GoTo EndProc
    Dim Orig_A() As Variant
    Orig_A = M
    A = M
    p = UBound(A, 1)
    ReDim Ematrix(1 To p, 1 To p)
    
    For iter = 1 To 500
        'Orthogonalize pairs of columns in upper off diag
        For i = 1 To p - 1
            For j = i + 1 To p

                x = GetColumn(A, i)
                y = GetColumn(A, j)
                num = 2 * WorksheetFunction.SumProduct(x, y)
                den = WorksheetFunction.SumSq(x) - WorksheetFunction.SumSq(y)
                
                'Skip rotation if aij is zero and correct ordering
                If Abs(num) < eps And den >= 0 Then Exit For
                
                'Perform Rotation
                If Abs(num) <= Abs(den) Then
                    Tan2 = Abs(num) / Abs(den)          ': eq. 11
                    Cos2 = 1 / Sqr(1 + Tan2 * Tan2)     ': eq. 12
                    Sin2 = Tan2 * Cos2                  ': eq. 13
                Else
                    Cot2 = Abs(den) / Abs(num)          ': eq. 16
                    Sin2 = 1 / Sqr(1 + Cot2 * Cot2)     ': eq. 17
                    Cos2 = Cot2 * Sin2                  ': eq. 18
                End If
                
                Cos_ = Sqr((1 + Cos2) / 2)              ': eq. 14/19
                Sin_ = Sin2 / (2 * Cos_)                ': eq. 15/20
                
                If den < 0 Then
                    tmp = Cos_
                    Cos_ = Sin_                         ': table 21
                    Sin_ = tmp
                End If
                
                Sin_ = Math.Sgn(num) * Sin_                  ': sign table 21
               
                'Rotate
                For k = 1 To p
                    A(k, i) = x(k) * Cos_ + y(k) * Sin_
                    A(k, j) = x(k) * -1 * Sin_ + y(k) * Cos_
                Next k
                
            Next j
        Next i
        
        'Test for convergence
        Test = Application.SumSq(A)
        If Abs(Test - hold) < eps And iter > 5 Then Exit For
        hold = Test
    Next iter
    
    If iter = 101 Then MsgBox "JK Iteration has not converged."
    Eval = WorksheetFunction.MMult(WorksheetFunction.MMult(WorksheetFunction.Transpose(A), Orig_A), A)
    'Compute eigenvalues/eigenvectors
    For i = 1 To p
        'Compute eigenvalues
        iSign = Math.Sgn(Eval(i, i))
        vEigenVal(i) = iSign * (iSign * Eval(i, i)) ^ (1 / 3)
        
        'Normalize eigenvectors
        iSign = Math.Sgn(A(1, i))
        For j = 1 To p
            Ematrix(j, i) = iSign * A(j, i) / Abs(vEigenVal(i))
        Next j
    Next i
    Set oEigenVec = New Matrix
    oEigenVec.Mat = Ematrix
    
    Exit Function
    
EndProc:
        MsgBox prompt:="Error in function PCA.Eigen()!" & vbCr & vbCr & _
            "Error: " & Err.Description & ".", Buttons:=48, _
            Title:="Run time error!"
End Function

Private Function GetColumn(A, c)
    Dim col As Variant
    ReDim col(1 To UBound(A))
    For i = 1 To UBound(A)
        col(i) = A(i, c)
    Next
    GetColumn = col
End Function

Public Sub Transform(iNum As Integer)
    Set oTransformed = oNormalData.Multiply(oEigenVec)
End Sub
