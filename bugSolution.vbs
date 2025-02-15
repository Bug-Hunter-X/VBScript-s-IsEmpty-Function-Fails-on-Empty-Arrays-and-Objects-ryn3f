Function f(a, b)
  If IsEmpty(a) Or IsEmpty(b) Then
    Exit Function
  End If
  If IsArray(a) Or IsArray(b) Or IsObject(a) Or IsObject(b) Then
    Err.Raise 13, , "Function f does not support array or object arguments."
  End If
  'Some code here that could cause an error
  If a > b Then
    c = a - b
  Else
    c = b - a
  End If
End Function