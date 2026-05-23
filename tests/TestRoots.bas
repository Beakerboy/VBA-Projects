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

Sub TestQuadratic()
    Result = Quadratic(1, 2, 1)
    Size = UBound(Result) - LBound(Result) + 1
    Debug.Assert Size = 2
    Root1 = Result(0)
    Size = UBound(Root1) - LBound(Root1) + 1
    Debug.Assert Size = 2
End Sub
