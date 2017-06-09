---
title: Series.PictureUnit2 Property (Excel)
keywords: vbaxl10.chm578123
f1_keywords:
- vbaxl10.chm578123
ms.prod: excel
api_name:
- Excel.Series.PictureUnit2
ms.assetid: 6c29fd60-2e42-3f7a-1fc0-67309ea75a3a
ms.date: 06/08/2017
---


# Series.PictureUnit2 Property (Excel)

Returns or sets the unit for each picture on the chart if the  **[PictureType](series-picturetype-property-excel.md)** property is set to **xlStackScale** (if not, this property is ignored). Read/write **Double** .


## Syntax

 _expression_ . **PictureUnit2**

 _expression_ A variable that represents a **Series** object.


## Example

This example sets series one in Chart1 to stack pictures and uses each picture to represent five units. The example should be run on a 2-D column chart with picture data markers.


```vb
With Charts("Chart1").SeriesCollection(1) 
 .PictureType = xlScale 
 .PictureUnit2 = 5 
End With
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

