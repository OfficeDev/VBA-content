---
title: PivotCell.PivotRowLine Property (Excel)
keywords: vbaxl10.chm692083
f1_keywords:
- vbaxl10.chm692083
ms.prod: excel
api_name:
- Excel.PivotCell.PivotRowLine
ms.assetid: e7e1ed02-b401-15b1-8548-fbdeb84796fc
ms.date: 06/08/2017
---


# PivotCell.PivotRowLine Property (Excel)

Returns the PivotLine on a row for a specific  **PivotCell** object. Read-only **PivotLine** .


## Syntax

 _expression_ . **PivotRowLine**

 _expression_ A variable that represents a **PivotCell** object.


## Remarks

If the PivotCell is on rows,  **PivotRowLine** returns the row's **PivotLine** object.

If the PivotCell is on columns,  **PivotRowLine** returns a run-time error.

If the PivotCell is in the data area, **PivotRowLine** returns the corresponding row's **PivotLine** object.


## See also


#### Concepts


[PivotCell Object](pivotcell-object-excel.md)

