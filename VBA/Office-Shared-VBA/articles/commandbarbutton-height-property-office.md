---
title: CommandBarButton.Height Property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Height
ms.assetid: b374ae8b-cce2-7562-1247-32ea90dc3c68
ms.date: 06/08/2017
---


# CommandBarButton.Height Property (Office)

Gets or sets the height of a command bar control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents a **CommandBarButton** object.


### Return Value

Integer


## Example

This example adds a custom control to the command bar named Custom. The example sets the height of the custom control to twice the height of the command bar and sets the control's width to 50 pixels. Notice how the command bar automatically resizes itself to accommodate the control.


```
Set myBar = CommandBars("Custom") 
barHeight = myBar.Height 
Set myControl = myBar.Controls _ 
    .Add(Type:=msoControlButton, _ 
    Id:= CommandBars("Standard").Controls("Save").Id, _ 
     Temporary:=True) 
With myControl 
    .Height = barHeight * 2 
    .Width = 50 
End With 
myBar.Visible = True
```


## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

