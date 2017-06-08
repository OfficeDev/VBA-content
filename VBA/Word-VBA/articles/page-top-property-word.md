---
title: Page.Top Property (Word)
keywords: vbawd10.chm11075587
f1_keywords:
- vbawd10.chm11075587
ms.prod: word
api_name:
- Word.Page.Top
ms.assetid: 01b3534c-fd22-720f-aa09-1f26f4fa335a
ms.date: 06/08/2017
---


# Page.Top Property (Word)

Returns a  **Long** that represents the top edge of the page. Read-only.


## Syntax

 _expression_ . **Top**

 _expression_ A variable that represents a **[Page](page-object-word.md)** object.


## Remarks

The  **Top** and **Left** properties of the **Page** object always return 0 (zero) indicating the upper left corner of the page. The **Height** and **Width** properties return the height and width in points (72 points = 1 inch) of the paper size specified in the Page Setup dialog or through the **PageSetup** object. For example, for an 8-1/2 by 11 inch page in portrait mode, the **Height** property returns 792 and the **Width** property returns 612. All four of these properties are read-only.


## See also


#### Concepts


[Page Object](page-object-word.md)

