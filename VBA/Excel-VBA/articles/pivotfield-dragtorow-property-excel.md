---
title: PivotField.DragToRow Property (Excel)
keywords: vbaxl10.chm240105
f1_keywords:
- vbaxl10.chm240105
ms.prod: excel
api_name:
- Excel.PivotField.DragToRow
ms.assetid: f10da457-1190-6b9f-ecc1-b9916c7fb4c4
ms.date: 06/08/2017
---


# PivotField.DragToRow Property (Excel)

 **True** if the field can be dragged to the row position. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **DragToRow**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

For OLAP data sources, the value is  **False** for measure fields.


## Example

This example prevents the Year field in the first PivotTable report on worksheet one from being dragged to the row position.


```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Year").DragToRow = False
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

