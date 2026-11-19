Attribute VB_NAME = "TestRoots"
Sub TestNegativeCubeRoot()
    Debug.Assert CubeRoot(-1) = -1
End Sub

Sub TestPositiveCubeRoot()
    Debug.Assert CubeRoot(1) = 1
End Sub

Sub TestDiscriminant()
    Debug.Assert Discriminant(1, 2, 3) = -8
End Sub

Sub TestQuadraticReal()
    Result = Quadratic(1, -2, 1)
    Size = UBound(Result) - LBound(Result) + 1
    Debug.Assert Size = 2
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = 0
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = 0
End Sub

Sub TestQuadraticCmplx()
    Result = Quadratic(1, 0, 1)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, -1)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = -1
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, 1)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = 1
End Sub

Sub TestQuadraticLinear()
    Result = Quadratic(0, 1, -1)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 1"
    Debug.Assert Size = 1
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = 0
End Sub

Sub TestCubicQuadratic()
    Result = Cubic(0, 1, -2, 1)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (1, 0)"
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = 0
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (1, 0)"
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = 0
End Sub

Sub TestCubicReal()
    Result = Cubic(1, 3, 3, 1)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 3"
    Debug.Assert Size = 3
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = -1
    Debug.Assert Root(2) = 0
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = -1
    Debug.Assert Root(2) = 0
    Root = Result(3)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = -1
    Debug.Assert Root(2) = 0
End Sub

Sub TestCubicCmplx()
    Result = Cubic(1, 1, 1, 1)
    Size = UBound(Result) - LBound(Result) + 1
    Debug.Assert Size = 3
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = -1
    Debug.Assert Root(2) = 0
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, 1)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = -1
    Root = Result(3)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, -1)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = 1
End Sub

Sub TestCubicOneZeroTwoReal()
    Result = Cubic(1, 0, 1, 0)
    Size = UBound(Result) - LBound(Result) + 1
    Debug.Assert Size = 3
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, 0)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = 0
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = -1
    Root = Result(3)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (1, 0)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = 1
End Sub

Sub TestCubicThreeUniqueReal()
    Result = Cubic(1, 2, -1, -2)
    Size = UBound(Result) - LBound(Result) + 1
    Debug.Assert Size = 3
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (1, 0)"
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = 0
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-2, 0)"
    Debug.Assert Root(1) = -1.9999999999999996
    Debug.Assert Root(2) = 0
    Root = Result(3)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = -1.0000000000000009
    Debug.Assert Root(2) = 0
End Sub

Sub TestQuarticLinear()
    Result = Quartic(0, 0, 0, 1, -1)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 1"
    Debug.Assert Size = 1
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = 0
End Sub

Sub TestQuarticRealEqual()
    Result = Quartic(1, 4, 6, 4, 1)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 4"
    Debug.Assert Size = 4
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = -1
    Debug.Assert Root(2) = 0
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = -1
    Debug.Assert Root(2) = 0
    Root = Result(3)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = -1
    Debug.Assert Root(2) = 0
    Root = Result(4)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = -1
    Debug.Assert Root(2) = 0
End Sub

Sub TestQuarticOneZero()
    Result = Quartic(1, 2, -1, -2, 0)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 4"
    Debug.Assert Size = 4
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, 0)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = 0
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (1, 0)"
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = 0
    Root = Result(3)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-2, 0)"
    Debug.Assert Root(1) = -1.9999999999999996
    Debug.Assert Root(2) = 0
    Root = Result(4)
    Size = UBound(Root) - LBound(Root) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = -1.0000000000000009
    Debug.Assert Root(2) = 0
End Sub

Sub TestBiQuarticReal()
    Result = Quartic(1, 0, -5, 0, 4)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 4"
    Debug.Assert Size = 4
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (1, 0)"
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = 0
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 0)"
    Debug.Assert Root(1) = -1
    Debug.Assert Root(2) = 0
    Root = Result(3)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (2, 0)"
    Debug.Assert Root(1) = 2
    Debug.Assert Root(2) = 0
    Root = Result(4)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-2, 0)"
    Debug.Assert Root(1) = -2
    Debug.Assert Root(2) = 0
End Sub

Sub TestBiQuarticImaginary()
    Result = Quartic(1, 0, 5, 0, 4)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 4"
    Debug.Assert Size = 4
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, 2)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = 2
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, -2)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = -2
    Root = Result(3)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, 1)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = 1
    Root = Result(4)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, -1)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = -1
End Sub

Sub TestBiQuarticComplex()
    Result = Quartic(1, 0, 0, 0, 4)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 4"
    Debug.Assert Size = 4
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (1, -1)"
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = -1
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, 1)"
    Debug.Assert Root(1) = -1
    Debug.Assert Root(2) = 1
    Root = Result(3)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (1, 1)"
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = 1
    Root = Result(4)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-1, -1)"
    Debug.Assert Root(1) = -1
    Debug.Assert Root(2) = -1
End Sub

Sub TestQuarticComplex()
    Result = Quartic(1, 1, -1, 1, -2)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 4"
    Debug.Assert Size = 4
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (-2, 0)"
    Debug.Assert Root(1) = -1.9999999999999998
    Debug.Assert Root(2) = 0
    Root = Result(2)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (1, 0)"
    Debug.Assert Root(1) = 0.9999999999999998
    Debug.Assert Root(2) = 0
    Root = Result(3)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, -1)"
    Debug.Assert Root(1) = 5.551115123125783e-17
    Debug.Assert Root(2) = -0.9999999999999998
    Root = Result(4)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & Root(1) & ", " & Root(2) & ") Expected: (0, 1)"
    Debug.Assert Root(1) = 0
    Debug.Assert Root(2) = 1
End Sub
