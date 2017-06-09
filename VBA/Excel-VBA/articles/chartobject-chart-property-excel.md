---
title: ChartObject.Chart Property (Excel)
keywords: vbaxl10.chm494099
f1_keywords:
- vbaxl10.chm494099
ms.prod: excel
api_name:
- Excel.ChartObject.Chart
ms.assetid: 99adb730-fc7b-1033-03e0-aebc82d95814
ms.date: 06/08/2017
---


# ChartObject.Chart Property (Excel)

Returns a  **[Chart](chart-object-excel.md)** object that represents the chart contained in the object. Read-only.


## Syntax

 _expression_ . **Chart**

 _expression_ A variable that represents a **ChartObject** object.


## Example

This example adds a title to the first embedded chart on Sheet1.


```vb
With Worksheets("Sheet1").ChartObjects(1).Chart 
 .HasTitle = True 
 .ChartTitle.Text = "1995 Rainfall Totals by Month" 
End With
```


## See also


#### Concepts


[ChartObject Object](chartobject-object-excel.md)

