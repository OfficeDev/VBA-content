---
title: NavigationControl.Controls Property (Access)
keywords: vbaac10.chm11036
f1_keywords:
- vbaac10.chm11036
ms.prod: access
api_name:
- Access.NavigationControl.Controls
ms.assetid: 68c6abcf-7bb7-4795-8c6c-685ed1c25dc9
ms.date: 06/08/2017
---


# NavigationControl.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **NavigationControl** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[NavigationControl Object](navigationcontrol-object-access.md)

