---
title: SparklineGroup.Location Property (Excel)
keywords: vbaxl10.chm871076
f1_keywords:
- vbaxl10.chm871076
ms.prod: excel
api_name:
- Excel.SparklineGroup.Location
ms.assetid: 3548cc42-dbab-636f-0dcf-2f38ad4a2db5
ms.date: 06/08/2017
---


# SparklineGroup.Location Property (Excel)

Gets or sets the  **[Range](range-object-excel.md)** object that represents the location of the sparkline group. Read/write.


## Syntax

 _expression_ . **Location**

 _expression_ A variable that represents a **[SparklineGroup](sparklinegroup-object-excel.md)** object.


### Return Value

Range


## Remarks

The location for all sparklines in a sparkline group must be on the same sheet, but the source data for the sparkline group can be on a different sheet or workbook.

The size of the range that represents the  **Location** property must equal the number of rows or columns in the **[SourceData](sparklinegroup-sourcedata-property-excel.md)** property for the **SparklineGroup** object.

A continuous range associated with a sparkline group can only be one dimensional. If the range is not continuous, each cell must be specified individually.


 **Note**  Do not use the  **[Union](application-union-method-excel.md)** method to create a non-continuous range because the **Union** method returns a single range reference.


## See also


#### Concepts


[SparklineGroup Object](sparklinegroup-object-excel.md)

