---
title: CommandBarComboBox.List Property (Office)
keywords: vbaof11.chm8005
f1_keywords:
- vbaof11.chm8005
ms.prod: office
api_name:
- Office.CommandBarComboBox.List
ms.assetid: c90fae92-daab-1b08-6e85-8caae26d0b72
ms.date: 06/08/2017
---


# CommandBarComboBox.List Property (Office)

Gets or sets an item in the  **CommandBarComboBox** control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **List**( **_Index_** )

 _expression_ A variable that represents a **CommandBarComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Integer**| The list item to be set.|

## Remarks




 **Note**  This property is read-only for built-in combo box controls.


## Example

This example checks the fourth list item in the combo box control whose caption is "Stock Data" on the command bar named "Custom." If the item isn't "View News," the example displays a message advising the user that the combo box may be damaged and asks the user to reinstall the application.


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
If CommandBars("Custom").Controls(1).List(4) _ 
     > "View News" Then 
MsgBox ("Stock Data appears to be damaged." &amp; _ 
     " Please reinstall application.") 
End If
```


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

