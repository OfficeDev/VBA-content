---
title: XlAllocationMethod Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlAllocationMethod
ms.assetid: ade163bf-81d2-f633-323a-603b7c96e867
ms.date: 06/08/2017
---


# XlAllocationMethod Enumeration (Excel)

Specifies the method to use to allocate values when performing what-if analysis on a PivotTable report based on an OLAP data source.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlEqualAllocation**|1|Use equal allocation.|
| **xlWeightedAllocation**|2|Use weighted allocation.|

## Remarks

If the  **[AllocationMethod](pivottable-allocationmethod-property-excel.md)** property is set to **xlWeightedAllocation** , you can optionally specify the weight expression to use by setting the **[AllocationWeightExpression](pivottable-allocationweightexpression-property-excel.md)** property.


