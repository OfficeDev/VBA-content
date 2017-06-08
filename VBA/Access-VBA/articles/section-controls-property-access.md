---
title: Section.Controls Property (Access)
keywords: vbaac10.chm12189
f1_keywords:
- vbaac10.chm12189
ms.prod: access
api_name:
- Access.Section.Controls
ms.assetid: 9cc617fd-716e-8d1e-8c2c-3808c5be55bb
ms.date: 06/08/2017
---


# Section.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **Section** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[Section Object](section-object-access.md)

