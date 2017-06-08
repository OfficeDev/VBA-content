---
title: FormatCondition.ShortestBarValue Property (Access)
keywords: vbaac10.chm14531
f1_keywords:
- vbaac10.chm14531
ms.prod: access
api_name:
- Access.FormatCondition.ShortestBarValue
ms.assetid: b262c385-0c12-87cc-45cc-83a658a61510
ms.date: 06/08/2017
---


# FormatCondition.ShortestBarValue Property (Access)

Gets or sets a numeric expression that specifies the value of the shortest bar of a  **[FormatCondition](formatcondition-object-access.md)**. Read/write **String**.


## Syntax

 _expression_. **ShortestBarValue**

 _expression_ A variable that represents a **FormatCondition** object.


## Remarks

By default, the  **ShortestBarValue** contains a zero-length string ("").

If the value of the  **[ShortestBarLimit](formatcondition-shortestbarlimit-property-access.md)** property is **acAutomatic**, then the **ShortestBarValue** is ignored.


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-access.md)

