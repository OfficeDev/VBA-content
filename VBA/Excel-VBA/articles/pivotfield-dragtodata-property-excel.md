---
title: PivotField.DragToData Property (Excel)
keywords: vbaxl10.chm240106
f1_keywords:
- vbaxl10.chm240106
ms.prod: excel
api_name:
- Excel.PivotField.DragToData
ms.assetid: 3149f842-83de-7cd2-2f53-2d15164ee1af
ms.date: 06/08/2017
---


# PivotField.DragToData Property (Excel)

 **True** if the specified field can be dragged to the data position. The default value is **True** . Read/write **Boolean**


## Syntax

 _expression_ . **DragToData**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

For OLAP data sources, the value is  **False** for measure fields.


## Example

This example prevents the Year field from being dragged to the data position in the first PivotTable report on the first worksheet.


```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Year").DragToData = False
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

