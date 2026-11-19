Attribute VB_NAME = "TestEllipseMod"
Sub TestUnitCircle()
    Circ = UnitCircle(5)
    Size = UBound(Circ) - LBound(Circ) + 1
    vbatest_msg = "Received: " & Size & " Expected: 5"
    Debug.Assert Size = 2
    pt = Circ(1)
    Size = UBound(pt) - LBound(pt) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & pt(1) & ", " & pt(2) & ") Expected: (0, 0)"
    Debug.Assert pt(1) = 0
    Debug.Assert pt(2) = 0
    pt = Circ(2)
    Size = UBound(pt) - LBound(pt) + 1
    vbatest_msg = "Received: " & Size & " Expected: 2"
    Debug.Assert Size = 2
    vbatest_msg = "Received: (" & pt(1) & ", " & pt(2) & ") Expected: (0, 1)"
    Debug.Assert pt(1) = 0
    Debug.Assert pt(2) = 1
End Sub
