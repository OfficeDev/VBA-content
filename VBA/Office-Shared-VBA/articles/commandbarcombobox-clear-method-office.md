---
title: CommandBarComboBox.Clear Method (Office)
keywords: vbaof11.chm8002
f1_keywords:
- vbaof11.chm8002
ms.prod: office
api_name:
- Office.CommandBarComboBox.Clear
ms.assetid: f60afda8-5740-c6f6-7f3b-315dc95c45f8
ms.date: 06/08/2017
---


# CommandBarComboBox.Clear Method (Office)

Removes all list items from a command bar combo box control (a drop-down list box or a combo box).


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Clear**

 _expression_ An expression that returns a **CommandBarComboBox** object.


## Remarks




 **Note**  This method fails if it is applied to a built-in command bar control (a control that comes with Microsoft Office).


## Example

This example checks the number of items in the combo box control on a command bar named  **Custom Bar**. If there are fewer than three items in the list in the combo box, the example clears the list, adds a new first item to the list, and then displays this new item as the default for the combo box control.


```
Set myBar = CommandBars("Custom Bar") 
Set myControl = myBar.Controls _ 
    Type:=msoControlComboBox) 
With myControl 
    .AddItem "First Item", 1 
    .AddItem "Second Item", 2 
End With 
If myControl.ListCount < 3 Then 
    myControl.Clear 
    myControl.AddItem Text:="New Item", Index:=1 
End If
```


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

