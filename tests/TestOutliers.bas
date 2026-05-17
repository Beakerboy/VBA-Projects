Attribute VB_NAME = "TestOutliers"
Sub TestGrubbsScore()
    Data = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Result = GrubbsScore(Data)
    Expected = 1.4863010829205867
    vbatest_msg = "Received: " + Result + " Expected: " + Expected
    Debug.Assert Result = Expected
End Sub

Sub TestCriticalGrubbs()
    Result = CriticalGrubbs(.95, 10)
    Debug.Assert Result = 0.71
End Sub

Sub TestCriticalDixon()
    Result = CriticalDixonQ(5, .95)
    Debug.Assert Result = 0.71
End Sub
