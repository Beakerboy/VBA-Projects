Attribute VB_NAME = "TestOutliers"
Sub TestGruggsScore()
    Data = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Result = GrubbsScore(Data)
    Debug.Assert Result = 1.48630108292059
End Sub

Sub TestCriticalGrubbs()
    Result = CriticalDixonQ(5, .95)
    Debug.Assert Result = 0.71
End Sub
