Attribute VB_NAME = "TestRoots"
Sub TestCubeRoot()
    Debug.Assert CubeRoot(-1) = -1
End Sub

Sub TestCubeRoot()
    Debug.Assert CubeRoot(1) = 1
End Sub

Sub TestDiscriminant()
    Debug.Assert Discriminant(1, 2, 3) = -8
End Sub
