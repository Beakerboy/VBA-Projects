Attribute VB_NAME = "TestEllipseMod"
Sub TestUnitCircle()
    Result = UnitCircle(5)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 5"
    Debug.Assert Size = 5
End Sub
