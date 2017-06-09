---
title: SubForm.Visible Property (Access)
keywords: vbaac10.chm11930
f1_keywords:
- vbaac10.chm11930
ms.prod: access
api_name:
- Access.SubForm.Visible
ms.assetid: 6566c85c-f906-d6a1-3e44-8f6bedde1167
ms.date: 06/08/2017
---


# SubForm.Visible Property (Access)

Returns or sets whether the object is visible. Read/write  **Boolean**.


## Syntax

 _expression_. **Visible**

 _expression_ A variable that represents a **SubForm** object.


## Remarks

To hide an object when printing, use the  **DisplayWhen** property.

You can use the  **Visible** property to hide a control on a form or report by including the property in a macro or event procedure that runs when the **Current** event occurs. For example, you can show or hide a congratulatory message next to a salesperson's monthly sales total in a sales report, depending on the sales total.


## See also


#### Concepts


[SubForm Object](subform-object-access.md)

