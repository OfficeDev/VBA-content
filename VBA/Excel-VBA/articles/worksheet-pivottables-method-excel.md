---
title: Worksheet.PivotTables Method (Excel)
keywords: vbaxl10.chm175118
f1_keywords:
- vbaxl10.chm175118
ms.prod: excel
api_name:
- Excel.Worksheet.PivotTables
ms.assetid: b60944cd-827d-15dc-d49e-c739c237de15
ms.date: 06/08/2017
---


# Worksheet.PivotTables Method (Excel)

Returns an object that represents either a single PivotTable report (a  **[PivotTable](pivottable-object-excel.md)** object) or a collection of all the PivotTable reports (a **[PivotTables](pivottables-object-excel.md)** object) on a worksheet. Read-only.


## Syntax

 _expression_ . **PivotTables**( **_Index_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the report.|

### Return Value

Object


## Example

This example sets the Sum of 1994 field in the first PivotTable report on the active sheet to use the SUM function.


```vb
ActiveSheet.PivotTables("PivotTable1"). _ 
 PivotFields("Sum of 1994").Function = xlSum
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

