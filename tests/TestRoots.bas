Attribute VB_NAME = "TestRoots"
Sub TestCubeRoot()
    Result = CubeRoot(-1)
    'Debug.Assert (Result = -1)
Emd Sub

Sub TestDiscriminant()
    Result = Discriminant(1, 2, 3)
    Debug.Assert (Result = -8)
End Sub
