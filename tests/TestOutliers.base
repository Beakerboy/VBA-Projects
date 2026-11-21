Attribute VB_NAME = "TestOutliers"
Sub TestGrubbsScore()
    Data = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Result = GrubbsScore(Data)
    Expected = 1.4863010829205867
    vbatest_msg = "Received: " & Result & " Expected: " & Expected
    Debug.Assert Result = Expected
End Sub

Sub TestCriticalGrubbs()
    Result = CriticalGrubbs(10, .95)
    Expected = 1.2856344242159952
    vbatest_msg = "Received: " & Result & " Expected: " & Expected
    Debug.Assert Result = Expected
End Sub

Sub TestDixonScore()
    Data = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Result = DixonQScore(Data)
    Expected = .1111111111111111
    vbatest_msg = "Received: " & Result & " Expected: " & Expected
    Debug.Assert Result = Expected
End Sub

Sub TestCriticalDixon()
    Dim Result As Double
    Result = CriticalDixonQ(5, .95)
    Dim Expected As Double
    Expected = 0.71
    vbatest_msg = "Received: " & Result & " Expected: " & Expected
    Debug.Assert Result = Expected

    Result = CriticalDixonQ(5, .99)
    Expected = 0.821
    vbatest_msg = "Received: " & Result & " Expected: " & Expected
    Debug.Assert Result = Expected
End Sub
