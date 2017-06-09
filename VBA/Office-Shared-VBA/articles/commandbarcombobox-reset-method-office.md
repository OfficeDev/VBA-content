---
title: CommandBarComboBox.Reset Method (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.Reset
ms.assetid: 28609b13-8036-a956-095a-1a6a748f00ad
ms.date: 06/08/2017
---


# CommandBarComboBox.Reset Method (Office)

Resets a built-in command bar to its default configuration, or resets a built-in  **CommandBarComboBox** control to its original function and face.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Reset**

 _expression_ A variable that represents a **CommandBarComboBox** object.


## Remarks

Resetting a built-in control restores the actions originally intended for the control and resets each of the control's properties back to its original state. Resetting a built-in command bar removes custom controls and restores built-in controls.


## Example

This example customizes a command bar combo box. First, the combo box is reset to its default state. Then two line items are added to the combo box and various properties are set. 


```
Set combo = CommandBars("Custom").Controls(2) 
combo.Reset 
With combo 
    .AddItem "First Item", 1 
    .AddItem "Second Item", 2 
    .DropDownLines = 3 
    .DropDownWidth = 75 
    .ListIndex = 0 
End With 

```


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

