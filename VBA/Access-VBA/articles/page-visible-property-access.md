---
title: Page.Visible Property (Access)
keywords: vbaac10.chm12153
f1_keywords:
- vbaac10.chm12153
ms.prod: access
api_name:
- Access.Page.Visible
ms.assetid: d01a5c26-18ee-2533-38d7-98a7bb84a971
ms.date: 06/08/2017
---


# Page.Visible Property (Access)

Returns or sets whether the object is visible. Read/write  **Boolean**.


## Syntax

 _expression_. **Visible**

 _expression_ A variable that represents a **Page** object.


## Remarks

To hide an object when printing, use the  **DisplayWhen** property.

You can use the  **Visible** property to hide a control on a form or report by including the property in a macro or event procedure that runs when the **Current** event occurs. For example, you can show or hide a congratulatory message next to a salesperson's monthly sales total in a sales report, depending on the sales total.


## See also


#### Concepts


[Page Object](page-object-access.md)

