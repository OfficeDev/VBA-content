---
title: CommandBarControl.DescriptionText Property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.DescriptionText
ms.assetid: 4f7b8e0d-1f3a-f751-86a7-3378f21ecf3d
ms.date: 06/08/2017
---


# CommandBarControl.DescriptionText Property (Office)

Gets or sets the description for a command bar control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **DescriptionText**

 _expression_ A variable that represents a **CommandBarControl** object.


### Return Value

String


## Remarks

The description is not displayed to the user, but it can be useful for documenting the behavior of the control for other developers. 


## Example

This example adds a control to a custom command bar, including a description of the control's behavior.


```
Set myBar = CommandBars.Add("Custom", msoBarTop, , True) 
myBar.Visible = True  
Set myControl = myBar.Controls _ 
    .Add(Type:=msoControlButton, ID:= _ 
    CommandBars("Standard").Controls("Paste").ID) 
With myControl 
    .DescriptionText = "Pastes the contents of the Clipboard" 
    .Caption = "Paste" 
End With
```


## See also


#### Concepts


[CommandBarControl Object](commandbarcontrol-object-office.md)
#### Other resources


[CommandBarControl Object Members](commandbarcontrol-members-office.md)

