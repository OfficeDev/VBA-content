---
title: CommandBarComboBox.Index Property (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.Index
ms.assetid: a844b760-d165-02aa-41ad-0bc75c55d0ed
ms.date: 06/08/2017
---


# CommandBarComboBox.Index Property (Office)

Gets a  **Long** representing the index number for a **CommandBarComboBox** object in the collection. Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Index**

 _expression_ A variable that represents a **CommandBarComboBox** object.


### Return Value

Long


## Remarks

The position of the first command bar control is 1. Separators are not counted in the  **CommandBarControls** collection.


## Example

This example searches the command bar named "Custom2" for a control with an  **Id** value of 23. If such a control is found and the index number of the control is greater than 5, the control will be positioned as the first control on the command bar.


```
Set myBar = CommandBars("Custom2") 
Set ctrl1 = myBar.FindControl(Id:=23) 
If ctrl1.Index > 5 Then 
    ctrl1.Move before:=1 
End If
```


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

