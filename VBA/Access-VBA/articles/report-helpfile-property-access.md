---
title: Report.HelpFile Property (Access)
keywords: vbaac10.chm13717
f1_keywords:
- vbaac10.chm13717
ms.prod: access
api_name:
- Access.Report.HelpFile
ms.assetid: cfab3009-c8e1-5f56-020a-e0a972e0de50
ms.date: 06/08/2017
---


# Report.HelpFile Property (Access)

The name of a help file associated with a report. Read/write  **String**.


## Syntax

 _expression_. **HelpFile**

 _expression_ A variable that represents a **Report** object.


## Example

This example adds a custom command bar with a combo box that tracks stock data. The example also specifies the Help topic to be displayed for the combo box when the user presses SHIFT+F1.


```vb
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


[Report Object](report-object-access.md)

