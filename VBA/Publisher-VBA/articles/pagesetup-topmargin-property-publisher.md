---
title: PageSetup.TopMargin Property (Publisher)
keywords: vbapb10.chm6946837
f1_keywords:
- vbapb10.chm6946837
ms.prod: publisher
api_name:
- Publisher.PageSetup.TopMargin
ms.assetid: 4eee9b1e-6c76-7831-13bc-25926c3c0f10
ms.date: 06/08/2017
---


# PageSetup.TopMargin Property (Publisher)

Returns a  **Variant** that represents the distance between the top edge of the printer sheet and the top edge of the publication pages. Read-only.


## Syntax

 _expression_. **TopMargin**

 _expression_A variable that represents a  **PageSetup** object.


### Return Value

Variant


## Remarks

Numeric values are evaluated as points. String values can be in any unit supported by Microsoft Publisher (for example, "2.5 in"). The valid range of possible values is from zero to the difference between the height of the sheet and the height of the publication pages.

The  **TopMargin** property returns a value only when you print multiple pages on a single sheet of printer paper. If you attempt to use it in other circumstances, Microsoft Publisher will return nothing.


