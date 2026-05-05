Attribute VB_NAME = "TestRoots"
Sub TestDiscriminant()
    Result = Discriminant(1, 2, 3)
    Debug.Assert Result = -8
End Sub
