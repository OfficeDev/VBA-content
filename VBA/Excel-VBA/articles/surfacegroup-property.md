---
title: SurfaceGroup Property
keywords: vbagr10.chm65558
f1_keywords:
- vbagr10.chm65558
ms.prod: excel
api_name:
- Excel.SurfaceGroup
ms.assetid: f22bfac3-6c3c-0c82-8ca5-e167dd01e132
ms.date: 06/08/2017
---


# SurfaceGroup Property

Returns a ChartGroup object that represents the surface chart group of a 3-D chart. Read-only ChartGroup object.

 _expression_. **SurfaceGroup**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Example

This example sets the 3-D surface group to use a different color for each data marker. The example should be run on a 3-D chart.


```vb
myChart.SurfaceGroup.VaryByCategories = True
```


