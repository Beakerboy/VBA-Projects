Option Explicit

Sub GetPass()
    Const a = 65, b = 66, c = 32, d = 126
    Dim i#, j#, k#, l#, m#, n#, o#, p#, q#, r#, s#, t#

    With ActiveSheet
        If .ProtectContents Then
            On Error Resume Next
            For i = a To b
                For j = a To b
                    For k = a To b
                        For l = a To b
                            For m = a To b
                                For n = a To b
                                    For o = a To b
                                        For p = a To b
                                            For q = a To b
                                                For r = a To b
                                                    For s = a To b
                                                        For t = c To d
                                                            Pass = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & _
            Chr(n) & Chr(o) & Chr(p) & Chr(q) & Chr(r) & Chr(s) & Chr(t)
                                                            .Unprotect Pass
                                                        Next t
                                                    Next s
                                                Next r
                                            Next q
                                        Next p
                                    Next o
                                Next n
                            Next m
                        Next l
                    Next k
                Next j
            Next i
            MsgBox Pass
        End If
    End With
End Sub
