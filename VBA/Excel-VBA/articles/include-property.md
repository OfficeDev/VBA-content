---
title: Include Property
keywords: vbagr10.chm65701
f1_keywords:
- vbagr10.chm65701
ms.prod: excel
api_name:
- Excel.Include
ms.assetid: ed92c49d-88fc-7f44-15cf-0641032157b2
ms.date: 06/08/2017
---


# Include Property

True if the data in the specified row or column is included in the chart. Read/write Variant.

 _expression_. **Include**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Example

This example causes the data in the second row on the datasheet to be excluded from the chart.


```vb
With myChart.Application.DataSheet 
 .Rows(2).Include = False 
End With
```


