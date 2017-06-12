---
title: Page.Left Property (Word)
keywords: vbawd10.chm11075586
f1_keywords:
- vbawd10.chm11075586
ms.prod: word
api_name:
- Word.Page.Left
ms.assetid: 681390e2-f121-5652-2923-aa460db0da64
ms.date: 06/08/2017
---


# Page.Left Property (Word)

Returns a  **Long** that represents the left edge of the page. Read-only.


## Syntax

 _expression_ . **Left**

 _expression_ A variable that represents a **[Page](page-object-word.md)** object.


## Remarks

The  **Top** and **Left** properties of the **Page** object always return 0 (zero) indicating the upper left corner of the page. The **Height** and **Width** properties return the height and width in points (72 points = 1 inch) of the paper size specified in the Page Setup dialog or through the **PageSetup** object. For example, for an 8-1/2 by 11 inch page in portrait mode, the **Height** property returns 792 and the **Width** property returns 612. All four of these properties are read-only.


## See also


#### Concepts


[Page Object](page-object-word.md)

