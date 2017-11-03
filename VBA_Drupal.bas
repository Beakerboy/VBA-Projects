Public Function CustomQuery(stSQL, Optional return_column = False)
  Set cnt = New ADODB.Connection
  stCon = "DSN=[DSN];Uid=[user];Pwd=[PW];"
  cnt.ConnectionString = stCon
  Set rst = New ADODB.Recordset
  cnt.Open
  return_id = 0
  If return_column <> False Then
    If return_column = "NEWID" Then
      stSQL2 = "SET NOCOUNT ON;" & stSQL & ";" & "SELECT SCOPE_IDENTITY() as new_id;"
      'MsgBox stSQL2
      Set rst = cnt.Execute(stSQL2)
      return_id = rst("new_id")
    Else
      rst.Open stSQL, cnt
      return_id = rst(return_column)
    End If
  Else
    rst.Open stSQL, cnt
  End If
  If CBool(rst.State And adStateOpen) = True Then rst.Close
  Set rst = Nothing
  If CBool(cnt.State And adStateOpen) = True Then cnt.Close
  Set cnt = Nothing
  If return_column <> False Then CustomQuery = return_id
End Function
