' Class: Matrix
' Perform matrix algebra and provide matrix attributes.
'
' Version: 210331

' Variable: dA
' A 2D array of doubles
' dA(intM, intN)
Dim dA() As Variant

' Variable: intM
' The number of rows in the matrix
Dim intM As Integer

' Variable: intM
' The number of columns in the matrix
Dim intN As Integer 'columns

' Constructor: Class_Initialize
' Initializes the object.
Private Sub Class_Initialize()
End Sub

' Property: Mat
' Set the matrix data.
'
' Parameter:
'   vA - A 2D array of numerical data.
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

Public Function isDiagonal()
    If intM = intN Then
        isDiagonal = True
    Else
        isDiagonal = False
    End If
End Function

Public Function Add(oInput As Matrix) As Matrix
    If Me.M <> oInput.M Or Me.N <> oInput.N Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Addition requires compatible sized matricies"
    End If
    Dim i As Integer
    Dim vSum() As Variant
    
    ReDim vSum(1 To Me.M, 1 To Me.N)
    For i = 1 To Me.M
        Dim j As Integer
        For j = 1 To Me.N
            vSum(i, j) = dA(i, j) + oInput.getValue(i, j)
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vSum
    Set Add = oResult
End Function

Public Function Subtract(oInput As Matrix) As Matrix
    If Me.M <> oInput.M Or Me.N <> oInput.N Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Subtraction requires compatible sized matricies"
    End If
    Dim i As Integer
    Dim vSum() As Variant
    
    ReDim vSum(1 To intM, 1 To intN)
    For i = 1 To Me.M
        Dim j As Integer
        For j = 1 To Me.N
            vSum(i, j) = dA(i, j) - oInput.getValue(i, j)
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vSum
    Set Subtract = oResult
End Function

Public Function Multiply(oInput As Matrix) As Matrix
    If Me.N <> oInput.M Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Multiplication requires compatible sized matricies"
    End If
    Dim vProduct() As Variant
    
    ReDim vProduct(1 To Me.M, 1 To oInput.N)
    Dim i As Long
    For i = 1 To Me.M
        Dim oRow As Vector
        Set oRow = Me.GetRow(i)
        Dim j As Long
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

Public Function ScalarMultiply(dInput) As Matrix
    Dim i As Integer
    Dim vProduct() As Variant
    
    ReDim vProduct(1 To Me.M, 1 To Me.N)
    For i = 1 To Me.M
        Dim j As Integer
        For j = 1 To Me.N
            vProduct(i, j) = dA(i, j) * dInput
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vProduct
    Set ScalarMultiply = oResult
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

' Function: Transpose
'
Public Function Transpose() As Matrix
    Dim vResult As Variant
    ReDim vResult(1 To intN, 1 To intM)
    For i = 1 To intM
        For j = 1 To intN
            vResult(j, i) = Me.getValue(i, j)
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vResult
    Set Transpose = oResult
End Function

Public Function GetRow(iRow) As Vector
    Dim vResult() As Variant
    ReDim vResult(1 To intN, 1 To 1)
    For i = 1 To intN
        vResult(i, 1) = dA(iRow, i)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set GetRow = oResult
End Function

Public Function GetColumn(iCol) As Vector
    Dim vResult() As Variant
    ReDim vResult(1 To intM, 1 To 1)
    For i = 1 To intM
        vResult(i, 1) = dA(i, iCol)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set GetColumn = oResult
End Function

Public Function getValue(i, j)
    getValue = dA(i, j)
End Function

Public Function AugmentRight(inMatrix As Matrix)
    If Me.M <> inMatrix.M Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Multiplication requires compatible sized matricies"
    End If
    Dim vResult As Variant
    vResult = dA
    ReDim Preserve vResult(1 To Me.M, 1 To Me.N + inMatrix.N)
    Dim oResults As New Matrix
    oResults.Mat = vResult
    Dim i As Long
    For i = 1 To inMatrix.N
        Set oResults = oResults.ReplaceColumn(Me.N + i, inMatrix.GetColumn(i))
    Next i
    Set AugmentRight = oResults
End Function

Public Function AugmentBelow(inMatrix As Matrix)
    Set AugmentBelow = Me.Transpose().AugmentRight(inMatrix.Transpose()).Transpose()
End Function

Public Function ReplaceColumn(iNum, inColumn As Vector) As Matrix
    If Me.M <> inColumn.M Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Multiplication requires compatible sized elements"
    End If
    vData = dA
    Dim i As Long
    For i = 1 To Me.M
        vData(i, iNum) = inColumn.getValue(i)
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vData
    Set ReplaceColumn = oResult
End Function

Public Function ReplaceRow(iNum, inRow As Vector) As Matrix
    If Me.N <> inRow.N Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Multiplication requires compatible sized elements"
    End If
    vData = dA
    Dim i As Long
    For i = 1 To Me.N
        vData(iNum, i) = inRow.getValue(i)
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vData
    Set ReplaceRow = oResult
End Function

Public Function ExcludeRow(iNum) As Matrix
    'Check that row exists
    Dim oResult As New Matrix
    If iNum = 1 Then
        Set oResult = Submatrix(2, intM, 1, intN)
    ElseIf iNum = intM Then
        Set oResult = Submatrix(1, intM - 1, 1, intN)
    Else
        Set oResult = Me.Submatrix(1, iNum - 1, 1, intN).AugmentBelow(Me.Submatrix(iNum + 1, intM, 1, intN))
    End If
    Set ExcludeRow = oResult
End Function

Public Property Get Trace()
    If intM <> intN Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Trace Requires a square Matrix"
    End If
    Trace = GetDiagonalElements.Sum
End Property

Public Function GetDiagonalElements() As Vector
    Dim vResult() As Variant
    ReDim vResult(1 To intM, 1 To 1)
    For i = 1 To intM
        vResult(i, 1) = dA(i, i)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set GetDiagonalElements = oResult
End Function

Public Function SwapColumns(col1, col2) As Matrix
    Dim oResult As Matrix
    Set oResult = Me.ReplaceColumn(col1, Me.GetColumn(col2)).ReplaceColumn(col2, Me.GetColumn(col1))
    
    Set SwapColumns = oResult
End Function

Public Function Submatrix(startm, endm, startn, endn)
    ' Check that ends are after starts and that no out of bounds
    Dim vResult As Variant
    ReDim vResult(1 To endm - startm + 1, 1 To endn - startn + 1)
    
    For i = startm To endm
        For j = startn To endn
            vResult(i - startm + 1, j - startn + 1) = dA(i, j)
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vResult
    Set Submatrix = oResult
End Function

Public Function Inverse() As Matrix
    Dim vResult As Variant
    vResult = ScalarMatrix(0, intM, intM).Mat
    If isDiagonal() Then
        For i = 1 To intM
            vResult(i, i) = 1 / dA(i, i)
        Next i
    End If
    Dim oResult As New Matrix
    oResult.Mat = vResult
    Set Inverse = oResult
End Function
