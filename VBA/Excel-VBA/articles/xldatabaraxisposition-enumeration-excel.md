---
title: XlDataBarAxisPosition Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlDataBarAxisPosition
ms.assetid: 5e447cc5-0bd1-c96a-2e3b-5d701489e61f
ms.date: 06/08/2017
---


# XlDataBarAxisPosition Enumeration (Excel)

Specifies the axis position for a range of cells with conditional formatting as data bars.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlDataBarAxisAutomatic**|0|Display the axis at a variable position based on the ratio of the minimum negative value to the maximum positive value in the range. Positive values are displayed in a left-to-right direction. Negative values are displayed in a right-to-left direction. When all values are positive or all values are negative, no axis is displayed.|
| **xlDataBarAxisMidpoint**|1|Display the axis at the midpoint of the cell regardless of the set of values in the range. Positive values are displayed in a left-to-right direction. Negative values are displayed in a right-to-left direction. |
| **xlDataBarAxisNone**|2|No axis is displayed, and both positive and negative values are displayed in the left-to-right direction.|

