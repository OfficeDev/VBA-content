
# Application.Ready Property (Excel)

 **Last modified:** July 28, 2015

Returns  **True** when the Microsoft Excel application is ready; **False** when the Excel application is not ready. Read-only **Boolean**.

## Syntax

 _expression_. **Ready**

 _expression_A variable that represents an  **Application** object.


## Example

In this example, Microsoft Excel checks to see if the  **Ready** property is set to **True**, and if so, a message displays "Application is ready." Otherwise, Excel displays the message "Application is not ready."


```
Sub UseReady() 
 
 If Application.Ready = True Then 
 MsgBox "Application is ready." 
 Else 
 MsgBox "Application is not ready." 
 End If 
 
End Sub
```


## See also


#### Concepts


 [Application Object](19b73597-5cf9-4f56-8227-b5211f657f6f.md)
#### Other resources


 [Application Object Members](4cb9ca42-8d07-cc9c-2d80-4eb9a5921e1e.md)
