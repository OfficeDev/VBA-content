---
title: CommandButton.Controls Property (Access)
keywords: vbaac10.chm10445
f1_keywords:
- vbaac10.chm10445
ms.prod: access
api_name:
- Access.CommandButton.Controls
ms.assetid: 017d583d-671e-7d9b-bdae-d67a7d94b4a8
ms.date: 06/08/2017
---


# CommandButton.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

