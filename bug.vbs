Late Binding:  VBScript's late binding can lead to runtime errors that are difficult to track down during development.  If you're not explicitly declaring object variables and their types, errors might only surface when a specific object isn't available at runtime.  For example, trying to use a method on an object that doesn't support it will fail silently or throw a generic error.

Example:
```vbscript
Set objExcel = CreateObject("Excel.Application")
' ... some code ...
' Assume Excel is not installed or accessible

'This might not throw any error immediately
MsgBox objExcel.ActiveWorkbook.Name
```

Implicit Type Conversion: VBScript's loose typing can cause unexpected behavior. Implicit type conversions can lead to data loss or incorrect calculations.  For example, comparing a string to a number will often result in unexpected boolean results.