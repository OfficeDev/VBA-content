---
title: Report.Controls Property (Access)
keywords: vbaac10.chm13794
f1_keywords:
- vbaac10.chm13794
ms.prod: access
api_name:
- Access.Report.Controls
ms.assetid: ea1ad090-91ba-d2c8-2a42-83227068548f
ms.date: 06/08/2017
---


# Report.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **Report** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[Report Object](report-object-access.md)

