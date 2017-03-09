'From:
'http://stackoverflow.com/questions/3417028/ellipse-around-the-data-in-matlab

Function Ellipse(Data, sigma, Optional NumPoints = 11)

p = 2 * WorksheetFunction.Norm_S_Dist(sigma, True) - 1
chi = WorksheetFunction.ChiSq_Inv(p, 2)
Dim Averages As Variant
ReDim Averages(1 To 2)
Averages(1) = WorksheetFunction.Average(Application.Index(Data, , 1))
Averages(2) = WorksheetFunction.Average(Application.Index(Data, , 2))
'Calculate yhe covarience matrix of the data
Covarience = CovarienceMatrix(Data)

'Scale the data by the confidence interval
Cov = bsxfun("multiply", Covarience, chi)

'Calculate our Eigenvalues
D = Eigenvalues(Cov)

V = Eigenvectors(Cov)

'extract diagonal
DV = ExtractDiagonal(D)

'Sort eigenvalues large to small
'If they have to be reordered, reorder eigenvectors as well
If (DV(2) > DV(1)) Then
  D1 = DV(1)
  DV(1) = DV(2)
  DV(2) = D1
  V11 = V(1, 1)
  V21 = V(2, 1)
  V(1, 1) = V(1, 2)
  V(2, 1) = V(2, 2)
  V(1, 2) = V11
  V(2, 2) = V21
End If

'Take the square root of Eigenvalues
DV(1) = Math.Sqr(DV(1))
DV(2) = Math.Sqr(DV(2))
'Convert back to a diagonal matrix
D = Diagonal(DV)

TranslationMatrix = WorksheetFunction.MMult(V, D)

Unit_Circle = WorksheetFunction.Transpose(UnitCircle(NumPoints))
CenteredEllipse = WorksheetFunction.Transpose(WorksheetFunction.MMult(TranslationMatrix, Unit_Circle))
Ellipse = bsxfun3("plus", CenteredEllipse, Averages)
End Function

'A diagonal Matrix of eigenvalues
Function Eigenvalues(Covarience)

a = 1
b = -1 * Covarience(1, 1) - Covarience(2, 2)
c = -1 * Covarience(2, 1) * Covarience(1, 2) + Covarience(1, 1) * Covarience(2, 2)

Eigenvalue1 = (-b - Math.Sqr(b * b - 4 * a * c)) / 2 / a
Eigenvalue2 = (-1 * b + Math.Sqr(b * b - 4 * a * c)) / 2 / a

Eigen = ScalerMatrix(0, 2, 2)
Eigen(1, 1) = Eigenvalue1
Eigen(2, 2) = Eigenvalue2
Eigenvalues = Eigen
End Function

Function Eigenvectors(Covarience)
Dim ReturnMatrix As Variant
ReDim ReturnMatrix(1 To 2, 1 To 2)

D = Eigenvalues(Covarience)
Lambda = bsxfun("multiply", Identity(2), D(2, 2))
AMinusLambda = bsxfun2("minus", Covarience, bsxfun("multiply", Identity(2), D(2, 2)))
A21 = AMinusLambda(1, 2)
A11 = AMinusLambda(1, 1)
ReturnMatrix(2, 1) = Math.Sqr(A21 ^ 2 / (A11 ^ 2 + A21 ^ 2))
ReturnMatrix(1, 2) = ReturnMatrix(2, 1)
ReturnMatrix(1, 1) = ReturnMatrix(2, 1) * A11 / A21
ReturnMatrix(2, 2) = -1 * ReturnMatrix(1, 1)
Eigenvectors = ReturnMatrix
End Function

'Create an Identity Matrix of a given size
Function Identity(Size_num)
ReturnMatrix = ScalerMatrix(0, Size_num, Size_num)
For i = 1 To Size_num
    ReturnMatrix(i, i) = 1
Next i
Identity = ReturnMatrix
End Function

'Sample Covarience Matrix
Function CovarienceMatrix(Data)
Y = Application.Index(Data, , 2)
Count = Application.WorksheetFunction.Count(Y)

OnesM = ScalerMatrix(1, Count, Count)

'Average the columns
PartA = Application.WorksheetFunction.MMult(OnesM, Data)
PartAScale = bsxfun("divide", PartA, Count)

'Center the data by subtracting the averages
PartB = bsxfun2("minus", Data, PartAScale)

PartBprime = Application.WorksheetFunction.Transpose(PartB)

PartC = Application.WorksheetFunction.MMult(PartBprime, PartB)
PartD = bsxfun("divide", PartC, Count - 1)
CovarienceMatrix = PartD
End Function

'Make a matrix of dimensions m x n where every element is the value 'value'
Function ScalerMatrix(value, matrix_length, matrix_width)
Dim matrix_object As Variant
ReDim matrix_object(1 To matrix_length, 1 To matrix_width)
For i = 1 To matrix_length
  For j = 1 To matrix_width
    matrix_object(i, j) = value
  Next j
Next i
ScalerMatrix = matrix_object
End Function

'apply an element-wise operation between a matrix and a scaler or two matricies
'currently works only on matricies of width=2 and unlimited length
Function bsxfun(operator_type, matrix_object, scaler_value)
  Dim ReturnMatrix As Variant
  ReDim ReturnMatrix(LBound(matrix_object, 1) To UBound(matrix_object, 1), 1 To 2)
  
  For i = LBound(matrix_object, 1) To UBound(matrix_object, 1)
    For j = 1 To 2
      If operator_type = "minus" Then
        ReturnMatrix(i, j) = matrix_object(i, j) - scaler_value
      ElseIf operator_type = "plus" Then
        ReturnMatrix(i, j) = matrix_object(i, j) + scaler_value
      ElseIf operator_type = "divide" Then
        ReturnMatrix(i, j) = matrix_object(i, j) / scaler_value
      ElseIf operator_type = "multiply" Then
        ReturnMatrix(i, j) = matrix_object(i, j) * scaler_value
      End If
    Next j
  Next i
  bsxfun = ReturnMatrix
End Function
'apply an element-wise operation between a matrix and a scaler or two matricies
'currently works only on matricies of width=2 and unlimited length
'We use the bounds on matrix_2 because matrix_object is sometimes a range,
'and lBound/UBound don't work on ranges
Function bsxfun2(operator_type, matrix_object, matrix_2)
  Dim ReturnMatrix As Variant
  ReDim ReturnMatrix(LBound(matrix_2, 1) To UBound(matrix_2, 1), 1 To 2)
  
  For i = LBound(matrix_2, 1) To UBound(matrix_2, 1)
    For j = 1 To 2
      'MsgBox matrix_object(i, j)
      operator_value = matrix_2(i, j)
      If operator_type = "minus" Then
        ReturnMatrix(i, j) = matrix_object(i, j) - operator_value
      ElseIf operator_type = "plus" Then
        ReturnMatrix(i, j) = matrix_object(i, j) + operator_value
      ElseIf operator_type = "divide" Then
        ReturnMatrix(i, j) = matrix_object(i, j) / operator_value
      ElseIf operator_type = "multiply" Then
        ReturnMatrix(i, j) = matrix_object(i, j) * operator_value
      End If
    Next j
  Next i
  bsxfun2 = ReturnMatrix
End Function

'apply an element-wise operation between a matrix and a vector
'currently works only on matricies of width=2 and unlimited length
Function bsxfun3(operator_type, matrix_object, scaler_value)
  Dim ReturnMatrix As Variant
  ReDim ReturnMatrix(LBound(matrix_object, 1) To UBound(matrix_object, 1), 1 To 2)
  
  For i = LBound(matrix_object, 1) To UBound(matrix_object, 1)
    For j = 1 To 2
      If operator_type = "minus" Then
        ReturnMatrix(i, j) = matrix_object(i, j) - scaler_value(j)
      ElseIf operator_type = "plus" Then
        ReturnMatrix(i, j) = matrix_object(i, j) + scaler_value(j)
      ElseIf operator_type = "divide" Then
        ReturnMatrix(i, j) = matrix_object(i, j) / scaler_value(j)
      ElseIf operator_type = "multiply" Then
        ReturnMatrix(i, j) = matrix_object(i, j) * scaler_value(j)
      End If
    Next j
  Next i
  bsxfun3 = ReturnMatrix
End Function

'Extract the diagonals from a matrix
Function ExtractDiagonal(matrix_object)
Dim ReturnVector As Variant
ReDim ReturnVector(LBound(matrix_object, 1) To UBound(matrix_object, 1))
For i = LBound(matrix_object, 1) To UBound(matrix_object, 1)
 ReturnVector(i) = matrix_object(i, i)
Next i
ExtractDiagonal = ReturnVector
End Function

'Create a diagonal matrix from a vector
Function Diagonal(VectorObject)
  ReturnMatrix = ScalerMatrix(0, UBound(VectorObject), UBound(VectorObject))
  For i = LBound(VectorObject) To UBound(VectorObject)
    ReturnMatrix(i, i) = VectorObject(i)
  Next i
  Diagonal = ReturnMatrix
End Function

'produce a unit circle with the number of points specified.
'the zero point is produced twice for ease of plotting
Function UnitCircle(NumPoints)
Dim ReturnVector As Variant
ReDim ReturnVector(1 To NumPoints, 1 To 2)
Pi = WorksheetFunction.Pi
For i = 1 To NumPoints
  ReturnVector(i, 1) = Math.Cos(2 * Pi * i / (NumPoints - 1))
  ReturnVector(i, 2) = Math.Sin(2 * Pi * i / (NumPoints - 1))
Next i
UnitCircle = ReturnVector
End Function
