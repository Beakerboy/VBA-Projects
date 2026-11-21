Attribute VB_NAME = "TestEllipseMod"
Sub TestUnitCircle()
    Result = UnitCircle(5)
    Size = UBound(Result) - LBound(Result) + 1
    vbatest_msg = "Received: " & Size & " Expected: 5"
    Debug.Assert Size = 5
    ptx = Result(1, 1)
    pty = Result(1, 2)
    vbatest_msg = "Received: (" & ptx & ", " & pty & ") Expected: (0, 1)"
    Debug.Assert ptx = 6.123233995736766e-17
    Debug.Assert pty = 1
    ptx = Result(2, 1)
    pty = Result(2, 2)
    vbatest_msg = "Received: (" & ptx & ", " & pty & ") Expected: (-1, 0)"
    Debug.Assert ptx = -1
    Debug.Assert pty = 1.2246467991473532e-16
End Sub
