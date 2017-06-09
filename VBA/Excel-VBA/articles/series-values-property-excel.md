---
title: Series.Values Property (Excel)
keywords: vbaxl10.chm578111
f1_keywords:
- vbaxl10.chm578111
ms.prod: excel
api_name:
- Excel.Series.Values
ms.assetid: 3db2577e-ef0e-75ea-412b-531d7e67c098
ms.date: 06/08/2017
---


# Series.Values Property (Excel)

Returns or sets a  **Variant** value that represents a collection of all the values in the series.


## Syntax

 _expression_ . **Values**

 _expression_ A variable that represents a **Series** object.


## Remarks

The value of this property can be a range on a worksheet or an array of constant values, but not a combination of both. See the examples for details.


## Example

This example sets the series values from a range.


```vb
Charts("Chart1").SeriesCollection(1).Values = _ 
 Worksheets("Sheet1").Range("C5:T5")
```

To assign a constant value to each individual data point, you must use an array.




```vb
Charts("Chart1").SeriesCollection(1).Values = _ 
 Array(1, 3, 5, 7, 11, 13, 17, 19)
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

