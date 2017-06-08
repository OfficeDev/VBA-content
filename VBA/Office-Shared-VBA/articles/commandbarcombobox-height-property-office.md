---
title: CommandBarComboBox.Height Property (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.Height
ms.assetid: a3afc8c0-1c35-acc0-905c-0af47e84827d
ms.date: 06/08/2017
---


# CommandBarComboBox.Height Property (Office)

Gets or sets the height of a  **CommandBarComboBox** control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents a **CommandBarComboBox** object.


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


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

