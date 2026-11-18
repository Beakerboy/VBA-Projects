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
    Result = Quadratic(1, 2, 1)
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
    Result = Quadratic(1, 1, 1)
    Size = UBound(Result) - LBound(Result) + 1
    Debug.Assert Size = 2
    Root = Result(1)
    Size = UBound(Root) - LBound(Root) + 1
    Debug.Assert Size = 2
    Debug.Assert Root(1) = 1
    Debug.Assert Root(2) = 1
End Sub
