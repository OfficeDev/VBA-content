---
title: PivotCell.PivotColumnLine Property (Excel)
keywords: vbaxl10.chm692084
f1_keywords:
- vbaxl10.chm692084
ms.prod: excel
api_name:
- Excel.PivotCell.PivotColumnLine
ms.assetid: 99d8e14e-28b5-4c0c-2f92-402fbb5c2ea8
ms.date: 06/08/2017
---


# PivotCell.PivotColumnLine Property (Excel)

Returns the  **PivotLine** on a column for a specific **PivotCell** object. Read-only **PivotLine** .


## Syntax

 _expression_ . **PivotColumnLine**

 _expression_ A variable that represents a **PivotCell** object.


## Remarks

If the PivotCell is on rows, the  **PivotColumnLine** property returns a run-time error.

If the PivotCell is on columns, the  **PivotColumnLine** property returns the column **PivotLine** object.

If the PivotCell is in the data area, the  **PivotColumnLine** property returns the corresponding column **PivotLine** object.


## See also


#### Concepts


[PivotCell Object](pivotcell-object-excel.md)

