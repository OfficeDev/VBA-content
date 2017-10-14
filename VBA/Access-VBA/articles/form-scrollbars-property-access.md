---
title: Form.ScrollBars Property (Access)
keywords: vbaac10.chm13363
f1_keywords:
- vbaac10.chm13363
ms.prod: access
api_name:
- Access.Form.ScrollBars
ms.assetid: d35e3e88-10ce-20f8-d4b1-305b27992395
ms.date: 06/08/2017
---


# Form.ScrollBars Property (Access)

You can use the  **ScrollBars** property to specify whether scroll bars appear on a form. Read/write **Byte**.


## Syntax

 _expression_. **ScrollBars**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **ScrollBars** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Neither |0| No scroll bars appear on the form.|
|Horizontal Only|1|Horizontal scroll bar appears on the form. |
|Vertical Only|2|Vertical scroll bar appears on the form.|
|Both|3|(Default) Vertical and horizontal scroll bars appear on the form. |
If your form is larger than the available display window, you can use the  **ScrollBars** property to allow the user to view the entire form.

You can use the  **NavigationButtons** property to control record navigation.


## See also


#### Concepts


[Form Object](form-object-access.md)

