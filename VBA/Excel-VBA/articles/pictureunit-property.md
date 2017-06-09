---
title: PictureUnit Property
keywords: vbagr10.chm3077575
f1_keywords:
- vbagr10.chm3077575
ms.prod: excel
api_name:
- Excel.PictureUnit
ms.assetid: 28a7cd8b-2558-87a1-158f-ff9a1dca8f41
ms.date: 06/08/2017
---


# PictureUnit Property

Returns or sets the unit for each picture on the chart if the PictureType property is set to xlScale (otherwise, this property is ignored). Read/write Long for all objects, except for the Walls object, which is read/write Variant.

 _expression_. **PictureUnit**

 _expression_ Required. An expression that returns one of the above objects.


## Example

This example sets series one to stack pictures and uses each picture to represent five units. The example should be run on a 2-D column chart with picture data markers.


```vb
With myChart.SeriesCollection(1) 
 .PictureType = xlScale 
 .PictureUnit = 5 
End With
```


