---
title: WdRulerStyle Enumeration (Word)
ms.prod: word
api_name:
- Word.WdRulerStyle
ms.assetid: 819d51d2-a097-b8bd-4788-55facf1de192
ms.date: 06/08/2017
---


# WdRulerStyle Enumeration (Word)

Specifies the way Word adjusts the table when the left indent is changed.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdAdjustFirstColumn**|2|Adjusts the left edge of the first column only, preserving the positions of the other columns and the right edge of the table.|
| **wdAdjustNone**|0|Adjusts the left edge of row or rows, preserving the width of all columns by shifting them to the left or right. This is the default value.|
| **wdAdjustProportional**|1|Adjusts the left edge of the first column, preserving the position of the right edge of the table by proportionally adjusting the widths of all the cells in the specified row or rows.|
| **wdAdjustSameWidth**|3|Adjusts the left edge of the first column, preserving the position of the right edge of the table by setting the widths of all the cells in the specified row or rows to the same value.|

