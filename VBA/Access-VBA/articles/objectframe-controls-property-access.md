---
title: ObjectFrame.Controls Property (Access)
keywords: vbaac10.chm11556
f1_keywords:
- vbaac10.chm11556
ms.prod: access
api_name:
- Access.ObjectFrame.Controls
ms.assetid: 42884347-14f3-0f0f-dc7e-3d2ae8154a49
ms.date: 06/08/2017
---


# ObjectFrame.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)

