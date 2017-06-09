---
title: WebBrowserControl.Controls Property (Access)
keywords: vbaac10.chm14355
f1_keywords:
- vbaac10.chm14355
ms.prod: access
api_name:
- Access.WebBrowserControl.Controls
ms.assetid: 864cbaf1-ad1c-7b74-1aac-3df61758c30e
ms.date: 06/08/2017
---


# WebBrowserControl.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **WebBrowserControl** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[WebBrowserControl Object](webbrowsercontrol-object-access.md)

