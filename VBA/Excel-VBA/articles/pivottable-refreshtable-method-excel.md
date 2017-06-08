---
title: PivotTable.RefreshTable Method (Excel)
keywords: vbaxl10.chm235092
f1_keywords:
- vbaxl10.chm235092
ms.prod: excel
api_name:
- Excel.PivotTable.RefreshTable
ms.assetid: 778743e3-c53a-23e3-73c6-c18339cd1ac2
ms.date: 06/08/2017
---


# PivotTable.RefreshTable Method (Excel)

Refreshes the PivotTable report from the source data. Returns  **True** if it's successful.


## Syntax

 _expression_ . **RefreshTable**

 _expression_ A variable that represents a **PivotTable** object.


### Return Value

Boolean


## Example

This example refreshes the PivotTable report.


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.RefreshTable
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

