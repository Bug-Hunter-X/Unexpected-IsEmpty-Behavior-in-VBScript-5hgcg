Function MyFunc(param1)
  If VarType(param1) = vbEmpty Or IsNull(param1) Or param1 = "" Then
    ' Handle empty, null, or zero-length string
    param1 = ""
  End If
  ' ... rest of the function
End Function