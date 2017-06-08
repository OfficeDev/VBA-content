---
title: PivotTable.RefreshName Property (Excel)
keywords: vbaxl10.chm235091
f1_keywords:
- vbaxl10.chm235091
ms.prod: excel
api_name:
- Excel.PivotTable.RefreshName
ms.assetid: 488d5e0c-61f9-0c85-ac1b-16dc98360bb4
ms.date: 06/08/2017
---


# PivotTable.RefreshName Property (Excel)

Returns the name of the person who last refreshed the PivotTable report data. Read-only  **String** .


## Syntax

 _expression_ . **RefreshName**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

For OLAP data sources, this property is updated after each query.


## Example

This example displays the name of the person who last refreshed the PivotTable report.


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
MsgBox "The data was last refreshed by " &; pvtTable.RefreshName
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

