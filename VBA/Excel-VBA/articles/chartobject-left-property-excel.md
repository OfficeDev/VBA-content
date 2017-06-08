---
title: ChartObject.Left Property (Excel)
keywords: vbaxl10.chm494084
f1_keywords:
- vbaxl10.chm494084
ms.prod: excel
api_name:
- Excel.ChartObject.Left
ms.assetid: 2b4964e2-624e-e53e-6efc-f792bf28a202
ms.date: 06/08/2017
---


# ChartObject.Left Property (Excel)

Returns or sets a  **Double** value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).


## Syntax

 _expression_ . **Left**

 _expression_ A variable that represents a **ChartObject** object.


## Example

This example aligns the left edge of the embedded chart with the left edge of column B.


```vb
With Worksheets("Sheet1") 
 .ChartObjects(1).Left = .Columns("B").Left 
End With
```


## See also


#### Concepts


[ChartObject Object](chartobject-object-excel.md)

