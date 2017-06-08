---
title: Page.Controls Property (Access)
keywords: vbaac10.chm12144
f1_keywords:
- vbaac10.chm12144
ms.prod: access
api_name:
- Access.Page.Controls
ms.assetid: 86f2f033-7622-7e5d-c727-a5c9b1b312e6
ms.date: 06/08/2017
---


# Page.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **Page** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[Page Object](page-object-access.md)

