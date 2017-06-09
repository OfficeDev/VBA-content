---
title: PivotTable.GrandTotalName Property (Excel)
keywords: vbaxl10.chm235133
f1_keywords:
- vbaxl10.chm235133
ms.prod: excel
api_name:
- Excel.PivotTable.GrandTotalName
ms.assetid: 7b0142aa-8b3d-a595-760e-b8ac5834e30f
ms.date: 06/08/2017
---


# PivotTable.GrandTotalName Property (Excel)

Returns or sets the text string label that is displayed in the grand total column or row heading in the specified PivotTable report. The default value is the string "Grand Total". Read/write  **String** .


## Syntax

 _expression_ . **GrandTotalName**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example sets the grand total heading label to "Regional Total" in the second PivotTable report on the active worksheet.


```vb
ActiveSheet.PivotTables("PivotTable2").GrandTotalName = "Regional Total"
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

