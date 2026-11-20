Attribute VB_NAME = "TestEllipseMod"
Sub TestUnitCircle()
    Result = UnitCircle(5)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 5"
    Debug.Assert Size = 5
    pt = Result(1)
    ptx = pt(1)
    pty = pt(2)
    vbatest_msg = "Received: (" & ptx & ", " & pty & ") Expected: (0, 0)"
    Assert ptx = 0
    Assert pty = 0
    pt = Result(2)
    ptx = pt(1)
    pty = pt(2)
    vbatest_msg = "Received: (" & ptx & ", " & pty & ") Expected: (0, 1)"
    Assert ptx = 0
    Assert pty = 1
End Sub
