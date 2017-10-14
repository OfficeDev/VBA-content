---
title: FormatCondition.LongestBarValue Property (Access)
keywords: vbaac10.chm14533
f1_keywords:
- vbaac10.chm14533
ms.prod: access
api_name:
- Access.FormatCondition.LongestBarValue
ms.assetid: bff378d6-138a-31bf-4587-d0f4ce827240
ms.date: 06/08/2017
---


# FormatCondition.LongestBarValue Property (Access)

Gets or sets a numeric expression that specifies the value of the longest bar of a  **[FormatCondition](formatcondition-object-access.md)**. Read/write **String**.


## Syntax

 _expression_. **LongestBarValue**

 _expression_ A variable that represents a **FormatCondition** object.


## Remarks

By default, the  **LongestBarValue** contains a zero-length string ("").

If the value of the  **[LongestBarLimit](formatcondition-longestbarlimit-property-access.md)** property is **acAutomatic**, then the **LongestBarValue** is ignored.


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-access.md)

