---
title: Series.BubbleSizes Property (Excel)
keywords: vbaxl10.chm578113
f1_keywords:
- vbaxl10.chm578113
ms.prod: excel
api_name:
- Excel.Series.BubbleSizes
ms.assetid: 41e56271-ec4c-7f9e-9642-174c8435e7d6
ms.date: 06/08/2017
---


# Series.BubbleSizes Property (Excel)

Returns or sets a string that refers to the worksheet cells containing the x-value, y-value and size data for the bubble chart. When you return the cell reference, it will return a string describing the cells in A1-style notation. To set the size data for the bubble chart, you must use R1C1-style notation. Applies only to bubble charts. Read/write  **Variant** .


## Syntax

 _expression_ . **BubbleSizes**

 _expression_ A variable that represents a **Series** object.


## Example

This example displays the cell reference for the cells that contain the bubble chart x-value, y-value and size data.


```vb
MsgBox Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).BubbleSizes
```

This example shows how to set this property using R1C1-style notation.




```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).BubbleSizes = "=Sheet1!r1c5:r5c5"
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

