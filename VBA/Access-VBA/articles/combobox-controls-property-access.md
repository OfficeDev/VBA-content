---
title: ComboBox.Controls Property (Access)
keywords: vbaac10.chm11367
f1_keywords:
- vbaac10.chm11367
ms.prod: access
api_name:
- Access.ComboBox.Controls
ms.assetid: 8f936303-1d90-d1cd-320f-de175df686cf
ms.date: 06/08/2017
---


# ComboBox.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

