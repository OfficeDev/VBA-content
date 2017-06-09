---
title: CommandBarComboBox.RemoveItem Method (Office)
keywords: vbaof11.chm8009
f1_keywords:
- vbaof11.chm8009
ms.prod: office
api_name:
- Office.CommandBarComboBox.RemoveItem
ms.assetid: 8a40dcca-c320-c27f-ae91-97c195d4f821
ms.date: 06/08/2017
---


# CommandBarComboBox.RemoveItem Method (Office)

Removes an item from a  **CommandBarComboBox** control.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **RemoveItem**( **_Index_** )

 _expression_ A variable that represents a **CommandBarComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Integer**|The item to be removed from the list.|

## Example

This example determines whether there are more than three items in the specified combo box. If there are more than three items, the example removes the second item, alters the style, and sets a new value. It also sets the  **Tag** property of the parent object (the CommandBarControl object) to show that the list has changed.


```
Set myBar = CommandBars _ 
    .Add(Name:="Custom", Position:=msoBarTop, _ 
    Temporary:=True) 
With myBar 
    .Controls.Add Type:=msoControlComboBox, ID:=1 
    .Visible = True  
End With 
With CommandBars("Custom").Controls(1) 
    .AddItem "Get Stock Quote", 1 
    .AddItem "View Chart", 2 
    .AddItem "View Fundamentals", 3 
    .AddItem "View News", 4 
    .Caption = "Stock Data" 
    .DescriptionText = "View Data For Stock" 
End With 
Set myControl = myBar.Controls(1) 
With myControl 
    If .ListCount > 3 Then 
        .RemoveItem 2 
        .Style = msoComboNormal 
        .Text = "New Default" 
         Set ctrl = .Parent 
    End If 
End With
```


 **Note**  


 **Note**  The property fails when applied to controls other than list controls.


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

