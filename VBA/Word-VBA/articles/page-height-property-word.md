---
title: Page.Height Property (Word)
keywords: vbawd10.chm11075589
f1_keywords:
- vbawd10.chm11075589
ms.prod: word
api_name:
- Word.Page.Height
ms.assetid: fe097fed-868b-cb09-f2ad-d53cda76a426
ms.date: 06/08/2017
---


# Page.Height Property (Word)

Returns a  **Long** that represents the height of a page, in pixels.


## Syntax

 _expression_ . **Top**

 _expression_ An expression that represents a **[Page](page-object-word.md)** object.


## Remarks

The  **[Top](page-top-property-word.md)** and **[Left](page-left-property-word.md)** properties of the **Page** object always return 0 (zero) indicating the upper left corner of the page. The **Height** and **[Width](page-width-property-word.md)** properties return the height and width in points (72 points = 1 inch) of the paper size specified in the **Page Setup** dialog box or through the **[PageSetup](pagesetup-object-word.md)** object. For example, for an 8-1/2 by 11 inch page in portrait mode, the **Height** property returns 792 and the **Width** property returns 612. All four of these properties are read-only.


## See also


#### Concepts


[Page Object](page-object-word.md)

