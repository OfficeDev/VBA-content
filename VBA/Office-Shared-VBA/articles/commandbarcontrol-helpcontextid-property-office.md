---
title: CommandBarControl.HelpContextId Property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.HelpContextId
ms.assetid: 56f41107-92ad-7cb5-f522-7a338f0d8cf9
ms.date: 06/08/2017
---


# CommandBarControl.HelpContextId Property (Office)

Gets or sets the Help context Id number for the Help topic attached to the  **CommandBarControl**. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **HelpContextId**

 _expression_ A variable that represents a **CommandBarControl** object.


### Return Value

Integer


## Remarks

To use this property, you must also set the HelpFile property. Help topics respond to Shift+F1.


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


[CommandBarControl Object](commandbarcontrol-object-office.md)
#### Other resources


[CommandBarControl Object Members](commandbarcontrol-members-office.md)

