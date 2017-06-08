---
title: Page.Width Property (Word)
keywords: vbawd10.chm11075588
f1_keywords:
- vbawd10.chm11075588
ms.prod: word
api_name:
- Word.Page.Width
ms.assetid: 530e4e99-4962-5887-6a1d-da328f43ffb8
ms.date: 06/08/2017
---


# Page.Width Property (Word)

Returns a  **Long** that represents the width, in points, of the paper defined in the **Page Setup** dialog box. Read-only **Long** .


## Syntax

 _expression_ . **Width**

 _expression_ A variable that represents a **[Page](page-object-word.md)** object.


## Remarks

The  **Top** and **Left** properties of the **Page** object always return 0 (zero) indicating the upper left corner of the page. The **Height** and **Width** properties return the height and width in points (72 points = 1 inch) of the paper size specified in the Page Setup dialog or through the **PageSetup** object. For example, for an 8-1/2 by 11 inch page in portrait mode, the **Height** property returns 792 and the **Width** property returns 612. All four of these properties are read-only.


## See also


#### Concepts


[Page Object](page-object-word.md)

