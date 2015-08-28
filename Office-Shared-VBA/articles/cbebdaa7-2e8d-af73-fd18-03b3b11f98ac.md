
# CommandBars.DisableCustomize Property (Office)

 **Last modified:** July 28, 2015

Is  **True** if toolbar customization is disabled. Read/write.

## Syntax

 _expression_. **DisableCustomize**

 _expression_A variable that represents a  **CommandBars** object.


## Example

The following example switches the  **DisableCustomize** property on or off.


```
Sub ToggleCustomize() 
 With Application.CommandBars 
 If .DisableCustomize = True Then 
 .DisableCustomize = False 
 Else 
 .DisableCustomize = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


 [CommandBars Object](0e312e21-14ee-5055-d604-b66e61c53b47.md)
#### Other resources


 [CommandBars Object Members](c11db22d-b7bb-20a2-a455-e441cb8d5bc0.md)
