---
title: NavigationButton.Controls Property (Access)
keywords: vbaac10.chm10445
f1_keywords:
- vbaac10.chm10445
ms.prod: access
api_name:
- Access.NavigationButton.Controls
ms.assetid: 21ea22a5-72a5-3b98-468c-6f2baa1110cf
ms.date: 06/08/2017
---


# NavigationButton.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **NavigationButton** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[NavigationButton Object](navigationbutton-object-access.md)

