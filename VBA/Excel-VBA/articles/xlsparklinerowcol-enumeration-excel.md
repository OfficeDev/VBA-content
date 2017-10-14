---
title: XlSparklineRowCol Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlSparklineRowCol
ms.assetid: 1b978b0d-c2a9-3367-cdef-429f79d84882
ms.date: 06/08/2017
---


# XlSparklineRowCol Enumeration (Excel)

Specifies how to plot the sparkline when the data on which it is based is in a square-shaped range.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **SparklineColumnsSquare**|2|Plot the data by columns.|
| **SparklineNonSquare**|0|The sparkline is not bound to data in a square-shaped range.|
| **SparklineRowsSquare**|1|Plot the data by rows.|

## Remarks

The  **XlSparklineRowCol** enumeration is used by the **[PlotBy](http://msdn.microsoft.com/library/bec64068-b9de-d857-829f-4ce061ce7585%28Office.15%29.aspx)** property of the **[SparklineGroup](sparklinegroup-object-excel.md)** object to determine how to plot chart in a sparkline when data on which it based is in a square-shaped range, such as A1:B2.


