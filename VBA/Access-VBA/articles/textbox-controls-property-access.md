---
title: TextBox.Controls Property (Access)
keywords: vbaac10.chm11036
f1_keywords:
- vbaac10.chm11036
ms.prod: access
api_name:
- Access.TextBox.Controls
ms.assetid: 00d5dede-0583-9f0e-191a-28f91a0327b3
ms.date: 06/08/2017
---


# TextBox.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

