Function f(a, b)
  If IsEmpty(a) Then
    Exit Function
  End If
  If IsEmpty(b) Then
    Exit Function
  End If
  'Some code here that could cause an error
  If a > b Then
    c = a - b
  Else
    c = b - a
  End If
End Function