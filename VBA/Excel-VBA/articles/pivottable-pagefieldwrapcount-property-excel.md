---
title: PivotTable.PageFieldWrapCount Property (Excel)
keywords: vbaxl10.chm235121
f1_keywords:
- vbaxl10.chm235121
ms.prod: excel
api_name:
- Excel.PivotTable.PageFieldWrapCount
ms.assetid: 930bfe25-362e-f907-d593-6898db07f55b
ms.date: 06/08/2017
---


# PivotTable.PageFieldWrapCount Property (Excel)

Returns or sets the number of page fields in each column or row in the PivotTable report. Read/write  **Long** .


## Syntax

 _expression_ . **PageFieldWrapCount**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example causes the PivotTable report to draw three page fields in a row before starting a new row.


```vb
With Worksheets(1).PivotTables("Pivot1") 
 .PageFieldOrder = xlOverThenDown 
 .PageFieldWrapCount = 3 
End With
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

