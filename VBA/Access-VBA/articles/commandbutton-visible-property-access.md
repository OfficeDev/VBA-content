---
title: CommandButton.Visible Property (Access)
keywords: vbaac10.chm10462
f1_keywords:
- vbaac10.chm10462
ms.prod: access
api_name:
- Access.CommandButton.Visible
ms.assetid: 980c1f93-ae95-3481-5358-ad5362ffc9e8
ms.date: 06/08/2017
---


# CommandButton.Visible Property (Access)

Returns or sets whether the object is visible. Read/write  **Boolean**.


## Syntax

 _expression_. **Visible**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

To hide an object when printing, use the  **DisplayWhen** property.

You can use the  **Visible** property to hide a control on a form or report by including the property in a macro or event procedure that runs when the **Current** event occurs. For example, you can show or hide a congratulatory message next to a salesperson's monthly sales total in a sales report, depending on the sales total.


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

