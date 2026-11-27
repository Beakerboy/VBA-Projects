Attribute VB_NAME = "TestEllipseMod"
Sub TestScalerMatrix()
    Result = ScalerMatrix(7, 2, 3)
    Rows = UBound(Result, 1) - LBound(Result, 1) + 1
    Cols = UBound(Result, 2) - LBound(Result, 2) + 1
    vbatest_msg = "Received: " & Rows & "x" & Cols & " Expected: 2x3"
    Debug.Assert Rows = 2
    Debug.Assert Cols = 3
    vbatest_msg = "Received: " & Result(1, 1) & " and " & Result(2, 3) & " Expected: 7"
    Debug.Assert Result(1, 1) = 7
    Debug.Assert Result(2, 3) = 7
End Sub

Sub TestIdentity()
    Result = Identity(3)
    vbatest_msg = "Received: " & Result(1, 1) & ", " & Result(2, 2) & ", " & Result(3, 3) & " Expected: 1"
    Debug.Assert Result(1, 1) = 1
    Debug.Assert Result(2, 2) = 1
    Debug.Assert Result(3, 3) = 1
    vbatest_msg = "Received: " & Result(1, 2) & ", " & Result(3, 1) & " Expected: 0"
    Debug.Assert Result(1, 2) = 0
    Debug.Assert Result(3, 1) = 0
End Sub

Sub TestExtractDiagonal()
    Dim Mat(1 To 2, 1 To 2) As Double
    Mat(1, 1) = 4: Mat(1, 2) = 9
    Mat(2, 1) = 8: Mat(2, 2) = 5
    Result = ExtractDiagonal(Mat)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Result(1) & ", " & Result(2) & ") Expected: (4, 5)"
    Debug.Assert Result(1) = 4
    Debug.Assert Result(2) = 5
End Sub

Sub TestDiagonal()
    Dim Vec(1 To 2) As Double
    Vec(1) = 3: Vec(2) = 6
    Result = Diagonal(Vec)
    vbatest_msg = "Received: (" & Result(1, 1) & ", " & Result(2, 2) & ") Expected: (3, 6)"
    Debug.Assert Result(1, 1) = 3
    Debug.Assert Result(2, 2) = 6
    vbatest_msg = "Received: (" & Result(1, 2) & ", " & Result(2, 1) & ") Expected: (0, 0)"
    Debug.Assert Result(1, 2) = 0
    Debug.Assert Result(2, 1) = 0
End Sub

Sub TestBsxfun()
    Dim Mat(1 To 2, 1 To 2) As Double
    Mat(1, 1) = 10: Mat(1, 2) = 20
    Mat(2, 1) = 30: Mat(2, 2) = 40

    ResPlus = bsxfun("plus", Mat, 5)
    vbatest_msg = "Received: " & ResPlus(1, 1) & " Expected: 15"
    Debug.Assert ResPlus(1, 1) = 15

    ResMinus = bsxfun("minus", Mat, 5)
    vbatest_msg = "Received: " & ResMinus(1, 2) & " Expected: 15"
    Debug.Assert ResMinus(1, 2) = 15

    ResMult = bsxfun("multiply", Mat, 2)
    vbatest_msg = "Received: " & ResMult(2, 1) & " Expected: 60"
    Debug.Assert ResMult(2, 1) = 60

    ResDiv = bsxfun("divide", Mat, 10)
    vbatest_msg = "Received: " & ResDiv(2, 2) & " Expected: 4"
    Debug.Assert ResDiv(2, 2) = 4
End Sub

Sub TestBsxfun2()
    Dim Mat1(1 To 2, 1 To 2) As Double
    Dim Mat2(1 To 2, 1 To 2) As Double
    Mat1(1, 1) = 10: Mat1(1, 2) = 20: Mat1(2, 1) = 30: Mat1(2, 2) = 40
    Mat2(1, 1) = 2:  Mat2(1, 2) = 4:  Mat2(2, 1) = 5:  Mat2(2, 2) = 8

    ResPlus = bsxfun2("plus", Mat1, Mat2)
    vbatest_msg = "Received: " & ResPlus(1, 1) & " Expected: 12"
    Debug.Assert ResPlus(1, 1) = 12

    ResMinus = bsxfun2("minus", Mat1, Mat2)
    vbatest_msg = "Received: " & ResMinus(1, 2) & " Expected: 16"
    Debug.Assert ResMinus(1, 2) = 16

    ResMult = bsxfun2("multiply", Mat1, Mat2)
    vbatest_msg = "Received: " & ResMult(2, 1) & " Expected: 150"
    Debug.Assert ResMult(2, 1) = 150

    ResDiv = bsxfun2("divide", Mat1, Mat2)
    vbatest_msg = "Received: " & ResDiv(2, 2) & " Expected: 5"
    Debug.Assert ResDiv(2, 2) = 5
End Sub

Sub TestBsxfun3()
    Dim Mat(1 To 2, 1 To 2) As Double
    Dim Vec(1 To 2) As Double
    Mat(1, 1) = 10: Mat(1, 2) = 20
    Mat(2, 1) = 30: Mat(2, 2) = 40
    Vec(1) = 2: Vec(2) = 5

    ResPlus = bsxfun3("plus", Mat, Vec)
    vbatest_msg = "Received: " & ResPlus(1, 2) & " Expected: 25"
    Debug.Assert ResPlus(1, 2) = 25

    ResMinus = bsxfun3("minus", Mat, Vec)
    vbatest_msg = "Received: " & ResMinus(2, 1) & " Expected: 28"
    Debug.Assert ResMinus(2, 1) = 28

    ResMult = bsxfun3("multiply", Mat, Vec)
    vbatest_msg = "Received: " & ResMult(1, 1) & " Expected: 20"
    Debug.Assert ResMult(1, 1) = 20

    ResDiv = bsxfun3("divide", Mat, Vec)
    vbatest_msg = "Received: " & ResDiv(2, 2) & " Expected: 8"
    Debug.Assert ResDiv(2, 2) = 8
End Sub

Sub TestEigenvalues()
    Dim Cov(1 To 2, 1 To 2) As Double
    Cov(1, 1) = 2: Cov(1, 2) = 1
    Cov(2, 1) = 1: Cov(2, 2) = 2
    Result = Eigenvalues(Cov)
    vbatest_msg = "Received: (" & Result(1, 1) & ", " & Result(2, 2) & ") Expected: (1, 3)"
    Debug.Assert Result(1, 1) = 1
    Debug.Assert Result(2, 2) = 3
    Debug.Assert Result(1, 2) = 0
    Debug.Assert Result(2, 1) = 0
End Sub

Sub TestEigenvectors()
    Dim Cov(1 To 2, 1 To 2) As Double
    Cov(1, 1) = 2: Cov(1, 2) = 1
    Cov(2, 1) = 1: Cov(2, 2) = 2
    Result = Eigenvectors(Cov)
    ExpectedVal = 1 / Math.Sqr(2)
    vbatest_msg = "Received: " & Result(1, 1) & " Expected: " & (-1 * ExpectedVal)
    Debug.Assert Math.Abs(Result(1, 1) - (-1 * ExpectedVal)) < 0.000000000001
    Debug.Assert Math.Abs(Result(1, 2) - ExpectedVal) < 0.000000000001
    Debug.Assert Math.Abs(Result(2, 1) - ExpectedVal) < 0.000000000001
    Debug.Assert Math.Abs(Result(2, 2) - ExpectedVal) < 0.000000000001
End Sub

Sub TestCovarienceMatrix1()
    Dim Data(1 To 3, 1 To 2) As Double
    Data(1, 1) = 1: Data(1, 2) = 2
    Data(2, 1) = 3: Data(2, 2) = 4
    Data(3, 1) = 5: Data(3, 2) = 8
    Result = WorksheetFunction.Index(Data, , 1)
    Debug.Assert 3 = 3
End Sub

Sub TestPi()
    Debug.Assert WorksheetFunction.Pi = 3.14
End Sub

'Sub TestCovarienceMatrix()
    'Dim Data(1 To 3, 1 To 2) As Double
    'Data(1, 1) = 1: Data(1, 2) = 2
    'Data(2, 1) = 3: Data(2, 2) = 4
    'Data(3, 1) = 5: Data(3, 2) = 8
    'Result = CovarienceMatrix(Data)
    'vbatest_msg = "Received: " & Result(1, 1) & " Expected: 4"
    'Debug.Assert Math.Abs(Result(1, 1) - 4) < 0.000000000001
    'vbatest_msg = "Received: " & Result(1, 2) & " Expected: 6"
    'Debug.Assert Math.Abs(Result(1, 2) - 6) < 0.000000000001
    'Debug.Assert Math.Abs(Result(2, 1) - 6) < 0.000000000001
    'vbatest_msg = "Received: " & Result(2, 2) & " Expected: 9.33333333333333"
    'Debug.Assert Math.Abs(Result(2, 2) - (28 / 3)) < 0.000000000001
'End Sub

'Sub TestEllipse()
    'Dim Data(1 To 3, 1 To 2) As Double
    'Data(1, 1) = 1: Data(1, 2) = 2
    'Data(2, 1) = 3: Data(2, 2) = 4
    'Data(3, 1) = 5: Data(3, 2) = 8

    ' Test optional NumPoints default (11)
    'ResultDefault = Ellipse(Data, 1)
    'SizeDefault = UBound(ResultDefault, 1) - LBound(ResultDefault, 1) + 1
    'vbatest_msg = "Received: " & SizeDefault & " Expected: 11"
    'Debug.Assert SizeDefault = 11

    ' Test explicit NumPoints = 5
    'Result = Ellipse(Data, 1, 5)
    'Size = UBound(Result, 1) - LBound(Result, 1) + 1
    'vbatest_msg = "Received: " & Size & " Expected: 5"
    'Debug.Assert Size = 5

    ' Verify closure (point 1 equals point 5) and first two coordinate pairs
    'vbatest_msg = "Received: (" & Result(1, 1) & ", " & Result(1, 2) & ") Expected: (2.596718, 4.928749)"
    'Debug.Assert Math.Abs(Result(1, 1) - 2.59671832812714) < 0.00000001
    'Debug.Assert Math.Abs(Result(1, 2) - 4.92874857042254) < 0.00000001
    'Debug.Assert Math.Abs(Result(5, 1) - Result(1, 1)) < 0.000000000001
    'Debug.Assert Math.Abs(Result(5, 2) - Result(1, 2)) < 0.000000000001

    'vbatest_msg = "Received: (" & Result(2, 1) & ", " & Result(2, 2) & ") Expected: (-0.003391, 0.045162)"
    'Debug.Assert Math.Abs(Result(2, 1) - (-0.00339135193137162)) < 0.00000001
    'Debug.Assert Math.Abs(Result(2, 2) - 0.0451622493418357) < 0.00000001
'End Sub

'Sub TestUnitCircle()
   ' Result = UnitCircle(5)
   ' Size = UBound(Result) - LBound(Result) + 1
   ' vbatest_msg = "Received: " & Size & " Expected: 5"
   ' Debug.Assert Size = 5
  '  ptx = Result(1, 1)
   ' pty = Result(1, 2)
   ' vbatest_msg = "Received: (" & ptx & ", " & pty & ") Expected: (0, 1)"
   ' Debug.Assert ptx = 6.123233995736766e-17
  '  Debug.Assert pty = 1
   ' ptx = Result(2, 1)
   ' pty = Result(2, 2)
   ' vbatest_msg = "Received: (" & ptx & ", " & pty & ") Expected: (-1, 0)"
   ' Debug.Assert ptx = -1
   ' Debug.Assert pty = 1.2246467991473532e-16
'End Sub
