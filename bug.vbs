Function f(a)
  If IsEmpty(a) Then
    Exit Function
  End If
  f = a + 1
End Function

MsgBox f(Null)