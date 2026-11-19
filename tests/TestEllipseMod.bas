Attribute VB_NAME = "TestEllipseMod"
Sub TestUnitCircle()
    Result = UnitCircle(5)
    Size = UBound(Result, 1) - LBound(Result, 1) + 1
    vbatest_msg = "Received: " & Size & " Expected: 5"
    Debug.Assert Size = 5
    Size = UBound(Result, 2) - LBound(Result, 2) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    'pt = Result(1)
End Sub
