Attribute VB_Name = "Eigen"
Private Function JKeigen(R As Range)
    Dim Arr() As Variant ' declare an unallocated array.
    Arr = Range(R.Address) ' Arr is now an allocated array
    JKeigen = EIGEN_JK(Arr)
End Function

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
Function EIGEN_JK(M) As Variant

    Dim A() As Variant, Ematrix() As Double
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
    ReDim Ematrix(1 To p, 1 To p + 1)
    
    For iter = 1 To 500
        'Orthogonalize pairs of columns in upper off diag
        For i = 1 To p - 1
            For j = i + 1 To p

                x = getColumn(A, i)
                y = getColumn(A, j)
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
        Ematrix(i, 1) = iSign * (iSign * Eval(i, i)) ^ (1 / 3)
        
        'Normalize eigenvectors
        For j = 1 To p
            Ematrix(j, i + 1) = A(j, i) / Abs(Ematrix(i, 1))
        Next j
    Next i
    EIGEN_JK = Ematrix
    Exit Function
    
EndProc:
        MsgBox prompt:="Error in function EIGEN_JK!" & vbCr & vbCr & _
            "Error: " & Err.Description & ".", Buttons:=48, _
            Title:="Run time error!"
End Function

Private Function getColumn(A, c)
    Dim col As Variant
    ReDim col(1 To UBound(A))
    For i = 1 To UBound(A)
        col(i) = A(i, c)
    Next
    getColumn = col
End Function
