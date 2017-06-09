---
title: PivotFilter.Active Property (Excel)
keywords: vbaxl10.chm770078
f1_keywords:
- vbaxl10.chm770078
ms.prod: excel
api_name:
- Excel.PivotFilter.Active
ms.assetid: 9fdbab3b-96e1-d821-5dc3-77a8a02c850a
ms.date: 06/08/2017
---


# PivotFilter.Active Property (Excel)

Returns whether the specified PivotFilter is active. Read-only  **Boolean** .


## Syntax

 _expression_ . **Active**

 _expression_ A variable that represents a **PivotFilter** object.


## Remarks

This property returns **True** when the PivotField of the filter is in the PivotTable and the filter is evaluated when the PivotTable is updated. It returns **False** when the PivotField of the filter is not in the PivotTable and has no effect on the PivotTable calculation.


## See also


#### Concepts


[PivotFilter Object](pivotfilter-object-excel.md)

