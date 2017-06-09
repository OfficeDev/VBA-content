---
title: CheckBox.Controls Property (Access)
keywords: vbaac10.chm10690
f1_keywords:
- vbaac10.chm10690
ms.prod: access
api_name:
- Access.CheckBox.Controls
ms.assetid: 4003f288-678f-57a7-0be7-a57517f14188
ms.date: 06/08/2017
---


# CheckBox.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **CheckBox** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[CheckBox Object](checkbox-object-access.md)

