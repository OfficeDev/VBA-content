---
title: PivotField.DragToColumn Property (Excel)
keywords: vbaxl10.chm240102
f1_keywords:
- vbaxl10.chm240102
ms.prod: excel
api_name:
- Excel.PivotField.DragToColumn
ms.assetid: 1e3ce788-5484-2504-37bb-a08770871c98
ms.date: 06/08/2017
---


# PivotField.DragToColumn Property (Excel)

 **True** if the specified field can be dragged to the column position. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **DragToColumn**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

For OLAP data sources, the value is  **False** for measure fields.


## Example

This example prevents the Year field in the first PivotTable report on worksheet one from being dragged to the column position.


```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Year").DragToColumn = False
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

