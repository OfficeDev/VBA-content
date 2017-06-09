---
title: Series.BarShape Property (Excel)
keywords: vbaxl10.chm578114
f1_keywords:
- vbaxl10.chm578114
ms.prod: excel
api_name:
- Excel.Series.BarShape
ms.assetid: 27af7eea-6ad3-e906-c5f8-d9e29314b32d
ms.date: 06/08/2017
---


# Series.BarShape Property (Excel)

Returns or sets the shape used with the 3-D bar or column chart. Read/write  **[XlBarShape](xlbarshape-enumeration-excel.md)** .


## Syntax

 _expression_ . **BarShape**

 _expression_ A variable that represents a **Series** object.


## Example

This example sets the shape used with series one on chart one.


```vb
Charts(1).SeriesCollection(1).BarShape = xlConeToPoint
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

