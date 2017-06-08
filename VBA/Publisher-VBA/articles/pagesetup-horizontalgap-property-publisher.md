---
title: PageSetup.HorizontalGap Property (Publisher)
keywords: vbapb10.chm6946818
f1_keywords:
- vbapb10.chm6946818
ms.prod: publisher
api_name:
- Publisher.PageSetup.HorizontalGap
ms.assetid: e8ee51e0-59b3-8fb6-21f6-87d67a96dd66
ms.date: 06/08/2017
---


# PageSetup.HorizontalGap Property (Publisher)

Returns a  **Variant** that represents the distance between the right edge of one publication page and left edge of the next publication page in the same row when multiple pages are printed on one sheet of printer paper. Read-only.


## Syntax

 _expression_. **HorizontalGap**

 _expression_A variable that represents a  **PageSetup** object.


### Return Value

Variant


## Remarks

Numeric values are evaluated as points; string values can be in any unit supported by Microsoft Publisher (for example, "2.5 in"). The valid range of possible values is from zero to the difference between the sheet width and the page width.

This property applies only to publications where multiple pages will be printed on each printer sheet. Using this property for any other publication raises an error.


