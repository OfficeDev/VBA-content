---
title: PivotTable.AllocationValue Property (Excel)
keywords: vbaxl10.chm235188
f1_keywords:
- vbaxl10.chm235188
ms.prod: excel
api_name:
- Excel.PivotTable.AllocationValue
ms.assetid: c68351d8-2959-46db-1f43-ca1bc71e14fc
ms.date: 06/08/2017
---


# PivotTable.AllocationValue Property (Excel)

Returns or sets what value to allocate when performing what-if analysis on a PivotTable report based on an OLAP data source. Read/write


## Syntax

 _expression_ . **AllocationValue**

 _expression_ A variable that represents a **[PivotTable](pivottable-object-excel.md)** object.


### Return Value

 **[XlAllocationValue](xlallocationvalue-enumeration-excel.md)**


## Remarks

The  **AllocationValue** property corresponds to the **Value to Allocate** setting in the **What-If Analysis Settings** dialog box. The default setting is **xlAllocateValue** , which corresponds to the **The value entered divided by the number of allocations** setting.


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

