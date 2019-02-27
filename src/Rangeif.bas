'Return the subset of range2 where range1 meets the specified condition
'range1 - an excel range
'condition - the value to match
'function_name - the function to evaluate
'range2 - the array of values
Public Function RangeIfAdvanced(range1, condition, function_name, range2)
  Dim Result As Variant
  Num = 0
  RowNumbers = WorksheetFunction.CountA(range1)
  ReDim Result(1 To RowNumbers, 1 To 1)
  For i = 1 To RowNumbers
    CellValue = range1(i, 1)
    If function_name = "odd" Then
      evaluation = WorksheetFunction.Odd(CellValue)
    ElseIf function_name = "floor" Then
      evaluation = WorksheetFunction.Floor(CellValue, 1)
    End If

    If evaluation = condition Then
      Num = Num + 1
      Result(Num, 1) = range2(i, 1)
    End If
  Next i
  'Redim Preseve is not working
  Dim FinalResult As Variant
  ReDim FinalResult(1 To Num, 1 To 1)
  For j = 1 To Num
    FinalResult(j, 1) = Result(j, 1)
  Next j
  RangeIfAdvanced = FinalResult
End Function
