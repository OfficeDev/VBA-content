---
title: Series.PictureType Property (Excel)
keywords: vbaxl10.chm578101
f1_keywords:
- vbaxl10.chm578101
ms.prod: excel
api_name:
- Excel.Series.PictureType
ms.assetid: 098dac46-ec2d-ea2d-71e9-1094a5f0b23a
ms.date: 06/08/2017
---


# Series.PictureType Property (Excel)

Returns or sets a  **[XlChartPictureType](xlchartpicturetype-enumeration-excel.md)** value that represents the way pictures are displayed on a column or bar picture chart.


## Syntax

 _expression_ . **PictureType**

 _expression_ A variable that represents a **Series** object.


## Example

This example sets series one in Chart1 to stretch pictures. The example should be run on a 2-D column chart with picture data markers.


```vb
Charts("Chart1").SeriesCollection(1).PictureType = xlStretch
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

