---
title: PivotField.DataRange Property (Excel)
keywords: vbaxl10.chm240078
f1_keywords:
- vbaxl10.chm240078
ms.prod: excel
api_name:
- Excel.PivotField.DataRange
ms.assetid: 14d5e4c4-1acb-aa02-6694-28e358afc881
ms.date: 06/08/2017
---


# PivotField.DataRange Property (Excel)

Returns a  **[Range](range-object-excel.md)** object as shown in the following table. Read-only.


## Syntax

 _expression_ . **DataRange**

 _expression_ A variable that represents a **PivotField** object.


## Remarks





|**Object**|**Data range**|
|:-----|:-----|
|Data field|Data contained in the field|
|Row, column, or page field|Items in the field|
|Item|Data qualified by the item|

## Example

This example selects the PivotTable items in the field named "REGION."


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
Worksheets("Sheet1").Activate 
pvtTable.PivotFields("REGION").DataRange.Select
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

