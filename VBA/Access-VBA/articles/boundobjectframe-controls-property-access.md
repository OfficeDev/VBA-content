---
title: BoundObjectFrame.Controls Property (Access)
keywords: vbaac10.chm10905
f1_keywords:
- vbaac10.chm10905
ms.prod: access
api_name:
- Access.BoundObjectFrame.Controls
ms.assetid: 65113d53-fa59-ff69-c398-2ce42abd9e0b
ms.date: 06/08/2017
---


# BoundObjectFrame.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

