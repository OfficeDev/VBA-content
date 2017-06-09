---
title: PivotTable.ListFormulas Method (Excel)
keywords: vbaxl10.chm235111
f1_keywords:
- vbaxl10.chm235111
ms.prod: excel
api_name:
- Excel.PivotTable.ListFormulas
ms.assetid: 48e2ac3c-25c7-2e41-177a-97954569d3ee
ms.date: 06/08/2017
---


# PivotTable.ListFormulas Method (Excel)

Creates a list of calculated PivotTable items and fields on a separate worksheet.


## Syntax

 _expression_ . **ListFormulas**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

This method isn't available for OLAP data sources.


## Example

This example creates a list of calculated items and fields for the first PivotTable report on worksheet one.


```vb
Worksheets(1).PivotTables(1).ListFormulas
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

