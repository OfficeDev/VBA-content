---
title: CommandBarPopup.Parameter Property (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBarPopup.Parameter
ms.assetid: 3ad7783e-3afd-0019-1cf9-eae93992479b
---


# CommandBarPopup.Parameter Property (Office)

Gets or sets a string that an application can use to execute a command from a  **CommandBarPopup** control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Parameter**

 _expression_ A variable that represents a **CommandBarPopup** object.


### Return Value

String


## Remarks

If the specified parameter is set for a built-in control, the application can modify its default behavior if it can parse and use the new value. If the parameter is set for custom controls, you can use it to send information to Visual Basic procedures, or you can use it to hold information about the control (similar to a second Tag property value).


## Example

This example assigns a new parameter to a control and sets the focus to the new button.


```vb
Set myControl = CommandBars("Custom").Controls(4) 
With myControl 
    .Copy , 1 
    .Parameter = "2" 
    .SetFocus 
End With
```


## See also


#### Concepts


[CommandBarPopup Object](commandbarpopup-object-office.md)

