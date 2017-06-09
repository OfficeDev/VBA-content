---
title: CommandBarComboBox.AddItem Method (Office)
keywords: vbaof11.chm8001
f1_keywords:
- vbaof11.chm8001
ms.prod: office
api_name:
- Office.CommandBarComboBox.AddItem
ms.assetid: 66109c4e-a75b-ebca-99e8-b6848316a04f
ms.date: 06/08/2017
---


# CommandBarComboBox.AddItem Method (Office)

Adds a list item to the specified command bar combo box control. The combo box control must be a custom control and must be a drop-down list box or a combo box.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **AddItem**( **_Text_**, **_Index_** )

 _expression_ A variable that represents a **CommandBarComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required|**String**|The text added to the control.|
| _Index_|Optional|**Variant**|The position of the item in the list. If this argument is omitted, the item is added to the end of the list.|

## Example

This example adds a combo box control to a command bar. Two items are added to the control, and the number of line items and the width of the combo box are set.


```
Set myBar = CommandBars("Custom") 
Set myControl = myBar.Controls.Add(Type:=msoControlComboBox, Id:=1) 
With myControl 
    .AddItem "First Item", 1 
    .AddItem "Second Item", 2 
    .DropDownLines = 3 
    .DropDownWidth = 75 
    .ListHeaderCount = 0 
End With
```


 **Note**  


 **Note**  This method will fail if it's applied to an edit box or a built-in combo box control.


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

