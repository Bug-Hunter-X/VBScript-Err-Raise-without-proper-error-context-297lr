Function MyFunction(param1)
  On Error GoTo ErrorHandler
  If IsEmpty(param1) Then
    Err.Raise 9999, , "Parameter cannot be empty"
  End If
  ' ... rest of the function
  Exit Function
ErrorHandler:
  MsgBox "Error Number: " & Err.Number & "
Description: " & Err.Description, vbCritical
  Err.Clear
End Function