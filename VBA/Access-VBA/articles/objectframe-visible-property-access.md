---
title: ObjectFrame.Visible Property (Access)
keywords: vbaac10.chm11579
f1_keywords:
- vbaac10.chm11579
ms.prod: access
api_name:
- Access.ObjectFrame.Visible
ms.assetid: 2461bccb-44c6-82b4-93a0-9e4f8231cf53
ms.date: 06/08/2017
---


# ObjectFrame.Visible Property (Access)

Returns or sets whether the object is visible. Read/write  **Boolean**.


## Syntax

 _expression_. **Visible**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

To hide an object when printing, use the  **DisplayWhen** property.

You can use the  **Visible** property to hide a control on a form or report by including the property in a macro or event procedure that runs when the **Current** event occurs. For example, you can show or hide a congratulatory message next to a salesperson's monthly sales total in a sales report, depending on the sales total.


## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)

