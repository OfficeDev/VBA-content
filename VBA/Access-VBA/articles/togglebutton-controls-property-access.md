---
title: ToggleButton.Controls Property (Access)
keywords: vbaac10.chm11695
f1_keywords:
- vbaac10.chm11695
ms.prod: access
api_name:
- Access.ToggleButton.Controls
ms.assetid: 99ef9045-10c0-d059-ea6b-be70b9c12a7a
ms.date: 06/08/2017
---


# ToggleButton.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **ToggleButton** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[ToggleButton Object](togglebutton-object-access.md)

