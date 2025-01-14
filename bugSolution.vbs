Function f(a)
  If IsNull(a) Then
    Exit Function
  End If
  f = a + 1
End Function

MsgBox f(Null) ' Now handles Null correctly