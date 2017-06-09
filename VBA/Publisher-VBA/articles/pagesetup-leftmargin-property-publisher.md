---
title: PageSetup.LeftMargin Property (Publisher)
keywords: vbapb10.chm6946819
f1_keywords:
- vbapb10.chm6946819
ms.prod: publisher
api_name:
- Publisher.PageSetup.LeftMargin
ms.assetid: 19fbb72e-bb6e-18e9-28f3-c7e99b071bfb
ms.date: 06/08/2017
---


# PageSetup.LeftMargin Property (Publisher)

Returns a  **Variant** that represents the distance (in points) between the left edge of the printer sheet and the left edge of the publication pages when multiple pages are printed on a single sheet. Read-only.


## Syntax

 _expression_. **LeftMargin**

 _expression_A variable that represents a  **PageSetup** object.


### Return Value

Variant


## Remarks

Numeric values are evaluated as points. String values can be in any unit supported by Microsoft Publisher (for example, 2.5 in). The valid range of possible values is from zero to the difference between the sheet width and the page width.

The  **LeftMargin** property returns a value only when you print multiple pages on a single sheet of printer paper. If you try to use the **LeftMargin** property in other circumstances, Microsoft Publisher will return **Nothing**.


