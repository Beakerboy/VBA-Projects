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
