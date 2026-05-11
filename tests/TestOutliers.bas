Attribute VB_NAME = "TestOutliers"
Sub TestCriticalGrubbs()
    Result = CriticalDixonQ(5, .95)
    Assert Result = 0.71
End Sub
