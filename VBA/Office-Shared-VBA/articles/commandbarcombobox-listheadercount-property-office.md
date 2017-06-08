---
title: CommandBarComboBox.ListHeaderCount Property (Office)
keywords: vbaof11.chm8007
f1_keywords:
- vbaof11.chm8007
ms.prod: office
api_name:
- Office.CommandBarComboBox.ListHeaderCount
ms.assetid: 54625ef5-2e09-5a39-7909-e775c4e9e0c4
ms.date: 06/08/2017
---


# CommandBarComboBox.ListHeaderCount Property (Office)

Gets or sets the number of list items in a  **CommandBarComboBox** control that appears above the separator line. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **ListHeaderCount**

 _expression_ A variable that represents a **CommandBarComboBox** object.


## Remarks


 **Note**  This property is read-only for built-in combo box controls.

A  **ListHeaderCount** property value of - 1 indicates that there's no separator line in the combo box control.


## Example

This example adds a combo box control to the command bar named "Custom" and then adds two items to the combo box. The example uses the  **ListHeaderCount** property to display a separator line between First Item and Second Item in the combo box. The example also sets the number of line items, the width of the combo box, and an empty default for the combo box.


```
Set myBar = CommandBars("Custom") 
Set myControl = myBar.Controls.Add(Type:=msoControlComboBox) 
With myControl 
    .AddItem Text:="First Item", Index:=1 
    .AddItem Text:="Second Item", Index:=2 
    .DropDownLines = 3 
    .DropDownWidth = 75 
    .ListHeaderCount = 1 
End With
```


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

