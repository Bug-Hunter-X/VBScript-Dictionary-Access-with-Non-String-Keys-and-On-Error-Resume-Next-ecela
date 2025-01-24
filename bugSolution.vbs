Function GetValue(obj, prop)
  If VarType(prop) <> vbString Then
    Err.Raise vbObjectError + 1, , "Error: Property must be a string."
  End If
  On Error Resume Next
  GetValue = obj(prop)
  If Err.Number <> 0 Then
    Err.Clear
    GetValue = Null
  End If
End Function

Dim obj
Set obj = CreateObject("Scripting.Dictionary")
obj.Add "key", "value"

Dim value
value = GetValue(obj, "key")
WScript.Echo value ' Output: value

value = GetValue(obj, "nonexistent")
WScript.Echo value ' Output: Null

On Error GoTo ErrHandler
value = GetValue(obj, 123) 
WScript.Echo value
Exit Sub
ErrHandler:
  WScript.Echo "Error: " & Err.Description
End Sub
