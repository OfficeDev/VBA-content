---
title: OptionGroup.Controls Property (Access)
keywords: vbaac10.chm10817
f1_keywords:
- vbaac10.chm10817
ms.prod: access
api_name:
- Access.OptionGroup.Controls
ms.assetid: e252c2b0-ab71-ed95-da04-62cec990f63e
ms.date: 06/08/2017
---


# OptionGroup.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents an **OptionGroup** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[OptionGroup Object](optiongroup-object-access.md)

