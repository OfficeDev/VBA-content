---
title: ValueChange.PivotCell Property (Excel)
keywords: vbaxl10.chm889075
f1_keywords:
- vbaxl10.chm889075
ms.prod: excel
api_name:
- Excel.ValueChange.PivotCell
ms.assetid: 332859df-b643-cf9b-9c61-108f9324cee5
ms.date: 06/08/2017
---


# ValueChange.PivotCell Property (Excel)

Returns a  **[PivotCell](pivotcell-object-excel.md)** object that represents the cell (tuple) that was changed. Read-only


## Syntax

 _expression_ . **PivotCell**

 _expression_ A variable that represents a **[ValueChange](valuechange-object-excel.md)** object.


### Return Value

 **PivotCell**


## Remarks

When the value of the  **[VisibleInPivotTable](valuechange-visibleinpivottable-property-excel.md)** property of the specified **ValueChange** object is **True** , the **PivotCell** property returns a **PivotCell** object for the cell (tuple) that was changed. When the value of the **VisibleInPivotTable** property of the specified **ValueChange** object is **False** , the **PivotCell** property returns **NULL** .


