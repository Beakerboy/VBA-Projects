Attribute VB_NAME = "TestOutliers"
Sub TestGrubbsScore()
    Data = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Result = GrubbsScore(Data)
    Expected = 1.48630108292059
    vbatest_msg = "Received: " + Result + " Expected: " + Expected
    Debug.Assert Result = Expected
End Sub

Sub TestCriticalGrubbs()
    Result = CriticalDixonQ(5, .95)
    Debug.Assert Result = 0.71
End Sub

Sub TestPieces()
    Data = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Result = WorksheetFunction.Count(Data)
    Expected = 10
    vbatest_msg = "Received: " + Result + " Expected: " + Expected
    Debug.Assert Result = Expected
End Sub
