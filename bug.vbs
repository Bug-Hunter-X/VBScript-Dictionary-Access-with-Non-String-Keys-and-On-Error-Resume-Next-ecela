Function GetValue(obj, prop)
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

value = GetValue(obj, 123) 'Error occurs here
WScript.Echo value ' Output: Null (Incorrect)