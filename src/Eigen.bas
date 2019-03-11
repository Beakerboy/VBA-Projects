' Function: Eigen
' Computes the eigenvalues and eigenvectors for a real symmetric positive     
' definite matrix using the "JK Method". 
'
' Parameters:
'   M - The source matrix
'
' Returns:
'   The first column of the return matrix contains the eigenvalues and the
'   rest of the p+1 columns contain the eigenvectors.
'
' About: Literature Source
'   KAISER,H.F. (1972) "THE JK METHOD: A PROCEDURE FOR FINDING THE
'   EIGENVALUES OF A REAL SYMMETRIC MATRIX", The Computer Journal, VOL.15,
'   271-273.
Function EIGEN_JK(ByRef M As Variant) As Variant

    Dim A() As Variant, Ematrix() As Double
    Dim i As Long, j As Long, k As Long, iter As Long, p As Long
    Dim den As Double, hold As Double, Sin_ As Double, num As Double
    Dim Sin2 As Double, Cos2 As Double, Cos_ As Double, Test As Double
    Dim Tan2 As Double, Cot2 As Double, tmp As Double
    Const eps As Double = 1E-16
    
    On Error GoTo EndProc
    
    A = M
    p = UBound(A, 1)
    ReDim Ematrix(1 To p, 1 To p + 1)
    
    For iter = 1 To 15  
        'Orthogonalize pairs of columns in upper off diag
        For j = 1 To p - 1
            For k = j + 1 To p
                
                den = 0#
                num = 0#
                'Perform single plane rotation
                For i = 1 To p
                    num = num + 2 * A(i, j) * A(i, k)   ': numerator eq. 11
                    den = den + (A(i, j) + A(i, k)) * _
                        (A(i, j) - A(i, k))             ': denominator eq. 11
                Next i
                
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
                
                Sin_ = Sgn(num) * Sin_                  ': sign table 21
                
                'Rotate
                For i = 1 To p
                    tmp = A(i, j)
                    A(i, j) = tmp * Cos_ + A(i, k) * Sin_
                    A(i, k) = -tmp * Sin_ + A(i, k) * Cos_
                Next i
                
            Next k
        Next j
        
        'Test for convergence
        Test = Application.SumSq(A)
        If Abs(Test - hold) < eps And iter > 5 Then Exit For
        hold = Test
    Next iter
    
    If iter = 16 Then MsgBox "JK Iteration has not converged."
    
    'Compute eigenvalues/eigenvectors
    For j = 1 To p
        'Compute eigenvalues
        For k = 1 To p
            Ematrix(j, 1) = Ematrix(j, 1) + A(k, j) ^ 2
        Next k
        Ematrix(j, 1) = Sqr(Ematrix(j, 1))
        
        'Normalize eigenvectors
        For i = 1 To p
            If Ematrix(j, 1) <= 0 Then
                Ematrix(i, j + 1) = 0
            Else
                Ematrix(i, j + 1) = A(i, j) / Ematrix(j, 1)
            End If
        Next i
    Next j
        
    EIGEN_JK = Ematrix
    
    Exit Function
    
    EndProc:
        MsgBox prompt:="Error in function EIGEN_JK!" & vbCr & vbCr & _
            "Error: " & Err.Description & ".", Buttons:=48, _
            Title:="Run time error!"
End Function
