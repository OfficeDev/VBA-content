---
title: PivotItem.ShowDetail Property (Excel)
keywords: vbaxl10.chm246082
f1_keywords:
- vbaxl10.chm246082
ms.prod: excel
api_name:
- Excel.PivotItem.ShowDetail
ms.assetid: d79e3f27-ff7a-9bf2-313d-e1add3e386a7
ms.date: 06/08/2017
---


# PivotItem.ShowDetail Property (Excel)

 **True** if the outline is expanded for the specified range (so that the detail of the column or row is visible). The specified range must be a single summary column or row in an outline. Read/write **Variant** . For the **PivotItem** object (or the **Range** object if the range is in a PivotTable report), this property is set to **True** if the item is showing detail.


## Syntax

 _expression_ . **ShowDetail**

 _expression_ A variable that represents a **PivotItem** object.


## Remarks

This property isn't available for OLAP data sources.

If the specified range isn't in a PivotTable report, the following statements are true:


- The range must be in a single summary row or column.
    
- This property returns  **False** if _any_ of the children of the row or column are hidden.
    
- Setting this property to  **True** is equivalent to unhiding all the children of the summary row or column.
    
- Setting this property to  **False** is equivalent to hiding all the children of the summary row or column.
    
If the specified range is in a PivotTable report, it's possible to set this property for more than one cell at a time if the range is contiguous. This property can be returned only if the range is a single cell.


## See also


#### Concepts


[PivotItem Object](pivotitem-object-excel.md)

