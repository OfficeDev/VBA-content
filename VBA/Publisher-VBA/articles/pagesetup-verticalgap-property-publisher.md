---
title: PageSetup.VerticalGap Property (Publisher)
keywords: vbapb10.chm6946838
f1_keywords:
- vbapb10.chm6946838
ms.prod: publisher
api_name:
- Publisher.PageSetup.VerticalGap
ms.assetid: 191d66c4-d168-625a-47b7-028167a98af9
ms.date: 06/08/2017
---


# PageSetup.VerticalGap Property (Publisher)

Returns a  **Variant** that represents the distance (in points) between the bottom edge of one publication page and top edge of the publication page below it when more than one publication page is printed on a single printer page. Read-only.


## Syntax

 _expression_. **VerticalGap**

 _expression_A variable that represents a  **PageSetup** object.


### Return Value

Variant


## Remarks

You can use the  **VerticalGap** property when you want to print multiple pages on a single sheet of printer paper. If the page size, including the values for the **VerticalGap** and **HorizontalGap** properties, is greater than half the paper size, Microsoft Publisher displays an error.


