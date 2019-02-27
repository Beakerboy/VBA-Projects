'Given a column of data, calculate the Grubbs Score
'G = Max|Xi-Xbar| / sigma
Public Function GrubbsScore(Data)
    Average = WorksheetFunction.Average(Data)
    StDev = WorksheetFunction.StDev(Data)
    Number = WorksheetFunction.Count(Data)
    Max = 0
    Dim i As Integer
    For i = 1 To Data.Count
        Value = Data(i, 1)
        Absvalue = Value - Average
        If Absvalue < 0 Then
            Absvalue = Absvalue * -1
        End If
        If Value <> "" And Absvalue > Max Then
            Max = Absvalue
        End If
    Next i
    GrubbsScore = Max / StDev
End Function

Public Function CriticalGrubbs(N, alpha, Optional tails = 1)
    T = WorksheetFunction.T_Inv(alpha / tails / N, N - 2)
    CriticalGrubbs = (N - 1) / N ^ 0.5 * (T ^ 2 / (N - 2 + T ^ 2)) ^ 0.5
End Function

Public Function DixonQScore(Data)
    Dim num As Integer
    Dim fRange As Double
    Dim fGap As Double
    num = WorksheetFunction.Count(Data)
    fGap = WorksheetFunction.Max(WorksheetFunction.Max(Data) - WorksheetFunction.Small(Data, num - 1), WorksheetFunction.Small(Data, 2) - WorksheetFunction.Min(Data))
    fRange = WorksheetFunction.Max(Data) - WorksheetFunction.Min(Data)
    DixonQScore = fGap / fRange
End Function

Public Function CriticalDixonQ(N As Integer, alpha)
    Dim num As Integer
    num = 10
    Dim Q95 As Variant
    Dim Q99 As Variant
    Q95 = Array(1, 1, 1, 0.97, 0.829, 0.71, 0.625, 0.568, 0.526, 0.493, 0.466)
    Q99 = Array(1, 1, 1, 0.994, 0.926, 0.821, 0.74, 0.68, 0.634, 0.598, 0.568)
    If alpha = 0.95 Then
        CriticalDixonQ = Q95(WorksheetFunction.Min(N, num))
    ElseIf alpha = 0.99 Then
        CriticalDixonQ = Q99(WorksheetFunction.Min(N, num))
    End If
End Function
