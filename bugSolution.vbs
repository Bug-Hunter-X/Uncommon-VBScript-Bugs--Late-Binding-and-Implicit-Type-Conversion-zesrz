Early Binding and Explicit Type Conversion:

To address the late binding and implicit type conversion issues, the solutions involve:

1. **Early Binding:** Declare object variables with their specific type using the Dim statement.  This allows the compiler to check for object compatibility at compile time rather than runtime.  You'll need to add a reference to the relevant library (like the Excel library). 

2. **Explicit Type Conversion:** Use functions like CInt, CStr, CLng, etc., to explicitly convert data types before comparisons or calculations. This prevents implicit conversion errors.

3. **Robust Error Handling:** Incorporate On Error Resume Next or On Error GoTo to handle potential errors gracefully and prevent your script from crashing.

Example (Solution):
```vbscript
On Error Resume Next 'Handle potential errors gracefully
Dim objExcel As Object 'Declare variable type
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then 
  MsgBox "Error creating Excel object: " & Err.Description
  Exit Sub
End If

' ... some code ...
If Not objExcel Is Nothing Then ' Check before using the object
    MsgBox CStr(objExcel.ActiveWorkbook.Name) ' Explicit type conversion
    objExcel.Quit
    Set objExcel = Nothing
End If
```