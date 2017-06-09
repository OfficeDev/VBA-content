---
title: Series.XValues Property (Excel)
keywords: vbaxl10.chm578112
f1_keywords:
- vbaxl10.chm578112
ms.prod: excel
api_name:
- Excel.Series.XValues
ms.assetid: 63715a3c-9d2d-6213-ac99-2c583773b45a
ms.date: 06/08/2017
---


# Series.XValues Property (Excel)

Returns or sets an array of x values for a chart series. The  **XValues** property can be set to a range on a worksheet or to an array of values, but it cannot be a combination of both. Read/write **Variant** .


## Syntax

 _expression_ . **XValues**

 _expression_ A variable that represents a **Series** object.


## Remarks

For PivotChart reports, this property is read-only.


## Example

This example sets the x values for series one in Chart1 to the range B1:B5 on Sheet1.


```vb
Charts("Chart1").SeriesCollection(1).XValues = _ 
 Worksheets("Sheet1").Range("B1:B5")
```

This example uses an array to set values for the individual points in series one in Chart1.




```vb
Charts("Chart1").SeriesCollection(1).XValues = _ 
 Array(5.0, 6.3, 12.6, 28, 50)
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

