---
title: SparklineGroup.SourceData Property (Excel)
keywords: vbaxl10.chm871077
f1_keywords:
- vbaxl10.chm871077
ms.prod: excel
api_name:
- Excel.SparklineGroup.SourceData
ms.assetid: b55c67a5-2cf8-4a36-a8d5-c7653f13ceb3
ms.date: 06/08/2017
---


# SparklineGroup.SourceData Property (Excel)

Returns or sets the range that contains the source data for the sparkline group. Read/write.


## Syntax

 _expression_ . **SourceData**

 _expression_ A variable that represents a **[SparklineGroup](sparklinegroup-object-excel.md)** object.


### Return Value

String


## Remarks

The number of rows or columns in the source data range must equal the number of cells in the  **[Location](sparklinegroup-location-property-excel.md)** property range.

The data source range for a single sparkline in the sparkline group must be continuous.


## See also


#### Concepts


[SparklineGroup Object](sparklinegroup-object-excel.md)

