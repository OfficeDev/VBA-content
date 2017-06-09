---
title: CommandBarComboBox.Text Property (Office)
keywords: vbaof11.chm8011
f1_keywords:
- vbaof11.chm8011
ms.prod: office
api_name:
- Office.CommandBarComboBox.Text
ms.assetid: 91aa73ff-260c-c241-35d0-50bebbbaf190
ms.date: 06/08/2017
---


# CommandBarComboBox.Text Property (Office)

Gets or sets the text in the display or edit portion of the  **CommandBarComboBox** control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Text**

 _expression_ A variable that represents a **CommandBarComboBox** object.


### Return Value

String


## Example

This example creates a new command bar named "Custom" and adds to it a combo box that contains four list items. The example then uses the Text property to set Item 3 as the default list item.


```
Set myBar = CommandBars _ 
    .Add(Name:="Custom", Position:=msoBarTop, _ 
    Temporary:=True) 
With myBar 
    .Controls.Add Type:=msoControlComboBox, ID:=1 
    .Visible = True  
End With 
Set testComboBox = CommandBars("Custom").Controls(1) 
With testComboBox 
    .AddItem "Item 1", 1 
    .AddItem "Item 2", 2 
    .AddItem "Item 3", 3 
    .AddItem "Item 4", 4 
    .Text = "Item 3" 
End With
```


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

