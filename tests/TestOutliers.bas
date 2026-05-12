Attribute VB_NAME = "TestOutliers"
Sub TestCriticalGrubbs()
    Result = CriticalDixonQ(5, .95)
    Debug.Assert Result = 0.71
End Sub
