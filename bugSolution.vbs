Function GetObjectInVBScript(objName)
  On Error Resume Next
  Set obj = GetObject(objName)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = Nothing
  End If
  Set GetObjectInVBScript = obj
End Function

' Example usage:
Set myExcelApp = GetObjectInVBScript("Excel.Application")
If myExcelApp Is Nothing Then
  MsgBox "Excel is not running.", vbExclamation
Else
  MsgBox "Excel is running.", vbInformation
  myExcelApp.Quit
  Set myExcelApp = Nothing
End If