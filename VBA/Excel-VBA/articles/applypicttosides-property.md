---
title: ApplyPictToSides Property
keywords: vbagr10.chm67195
f1_keywords:
- vbagr10.chm67195
ms.prod: excel
api_name:
- Excel.ApplyPictToSides
ms.assetid: aa6146cf-4e4f-b0c7-55eb-0ed8bd9dcc65
ms.date: 06/08/2017
---


# ApplyPictToSides Property

 **True** if a picture is applied to the sides of the point or all points in the series. Read/write **Boolean**.


## Example

This example applies pictures to the sides of all points in series one. The series must already have pictures applied to it (setting this property changes the picture orientation).


```vb
myChart.SeriesCollection(1).ApplyPictToSides = True
```


