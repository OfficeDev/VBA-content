
# ODBCError.Application Property (Excel)

 **Last modified:** July 28, 2015

When used without an object qualifier, this property returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)**object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.

## Syntax

 _expression_. **Application**

 _expression_A variable that represents an  **ODBCError** object.


## Example

This example displays a message about the application that created  `myObject`.


```
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## See also


#### Concepts


 [ODBCError Object](a256d466-7fa1-4b0f-fe01-c2640743e7e9.md)
#### Other resources


 [ODBCError Object Members](d2dc90a0-5f7e-1e2e-6fdf-307b3ed42fec.md)
