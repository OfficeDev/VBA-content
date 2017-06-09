---
title: ChartObject.ZOrder Property (Excel)
keywords: vbaxl10.chm494096
f1_keywords:
- vbaxl10.chm494096
ms.prod: excel
api_name:
- Excel.ChartObject.ZOrder
ms.assetid: 1d3e3557-66c5-78f8-a86c-c0d64af63bc6
ms.date: 06/08/2017
---


# ChartObject.ZOrder Property (Excel)

Returns the z-order position of the object. Read-only  **Long** .


## Syntax

 _expression_ . **ZOrder**

 _expression_ A variable that represents a **ChartObject** object.


## Remarks

In any collection of objects, the object at the back of the z-order is  _collection_(1), and the object at the front of the z-order is  _collection_( _collection_. **Count** ). For example, if there are embedded charts on the active sheet, the chart at the back of the z-order is `ActiveSheet.ChartObjects(1)`, and the chart at the front of the z-order is  `ActiveSheet.ChartObjects(ActiveSheet.ChartObjects.Count)`.


## Example

This example displays the z-order position of embedded chart one on Sheet1.


```vb
MsgBox "The chart's z-order position is " &; _ 
 Worksheets("Sheet1").ChartObjects(1).ZOrder
```


## See also


#### Concepts


[ChartObject Object](chartobject-object-excel.md)

