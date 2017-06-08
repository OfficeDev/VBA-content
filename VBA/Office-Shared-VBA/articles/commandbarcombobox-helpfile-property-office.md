---
title: CommandBarComboBox.HelpFile Property (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.HelpFile
ms.assetid: 2fbe2d70-b8f7-d800-ed46-0ac88125b8f1
ms.date: 06/08/2017
---


# CommandBarComboBox.HelpFile Property (Office)

Gets or sets the file name for the Help topic attached to the  **CommandBarComboBox** control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **HelpFile**

 _expression_ A variable that represents a **CommandBarComboBox** object.


### Return Value

String


## Remarks

To use this property, you must also set the HelpContextID property. Help topics respond to the user pressing SHIFT+F1.


## Example

This example adds a custom command bar with a combo box that tracks stock data. The example also specifies the Help topic to be displayed for the combo box when the user presses SHIFT+F1.


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
    .HelpFile = "C:\corphelp\custom.hlp" 
    .HelpContextID = 47 
End With
```


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

