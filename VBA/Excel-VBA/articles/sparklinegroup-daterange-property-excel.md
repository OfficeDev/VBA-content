---
title: SparklineGroup.DateRange Property (Excel)
keywords: vbaxl10.chm871078
f1_keywords:
- vbaxl10.chm871078
ms.prod: excel
api_name:
- Excel.SparklineGroup.DateRange
ms.assetid: 4944aa78-89cc-8252-2c5e-148ca4229579
ms.date: 06/08/2017
---


# SparklineGroup.DateRange Property (Excel)

Gets or sets the date range for the sparkline group. Read/write.


## Syntax

 _expression_ . **DateRange**

 _expression_ A variable that represents a **[SparklineGroup](sparklinegroup-object-excel.md)** object.


### Return Value

String


## Remarks

To clear the date range set this property to an empty string.

The date range must be a continuous one dimensional range.

The date range can be located on a different sheet than the  **[Location](sparklinegroup-location-property-excel.md)** and **[SourceData](sparklinegroup-sourcedata-property-excel.md)** properties.

Empty cells and non-date values in the date range are not displayed.


## See also


#### Concepts


[SparklineGroup Object](sparklinegroup-object-excel.md)

