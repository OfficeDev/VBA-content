---
title: CommandBarButton.Reset Method (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Reset
ms.assetid: 0e39c960-3928-f91a-cf7e-1df5a2fd217b
ms.date: 06/08/2017
---


# CommandBarButton.Reset Method (Office)

Resets a built-in  **CommandBarButton** control to its original function and face.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Reset**

 _expression_ A variable that represents a **CommandBarButton** object.


## Remarks

Resetting a built-in control restores the actions originally intended for the control and resets each of the control's properties back to its original state.


## Example

This example customizes a command bar button. First, the button properties are reset to their default state. Then various button properties are set. 


```
Dim cbButton As CommandBarButton 
Set cbButton = CommandBars("Custom").Controls(2) 
cbButton.Reset 
With cbButton 
    .BuiltInFace = True  
    .Caption = "Compute Total" 
    .DescriptionText = "This button computes the total of all purchases." 
    .Enabled = True  
    .TooltipText = "Click to compute total amount for all items in your cart." 
End With
```


## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

