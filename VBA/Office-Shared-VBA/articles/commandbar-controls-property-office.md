---
title: CommandBar.Controls Property (Office)
keywords: vbaof11.chm3003
f1_keywords:
- vbaof11.chm3003
ms.prod: office
api_name:
- Office.CommandBar.Controls
ms.assetid: 5c025bc5-9266-18a2-21ee-6aee478fb322
ms.date: 06/08/2017
---


# CommandBar.Controls Property (Office)

Gets a  **CommandBarControls** object that represents all the controls on a command bar. Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **CommandBar** object.


### Return Value

CommandBarControls


## Example

This example adds a combo box control to the command bar named "Custom" and fills the list with two items. The example also sets the number of line items, the width of the combo box, and an empty default for the combo box.


```
Set myControl = CommandBars("Custom").Controls _ 
    .Add(Type:=msoControlComboBox, Before:=1) 
With myControl 
    .AddItem Text:="First Item", Index:=1 
    .AddItem Text:="Second Item", Index:=2 
    .DropDownLines = 3 
    .DropDownWidth = 75 
    .ListHeaderCount = 0 
End With
```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

