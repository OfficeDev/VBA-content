---
title: PictureType Property
keywords: vbagr10.chm3077574
f1_keywords:
- vbagr10.chm3077574
ms.prod: excel
api_name:
- Excel.PictureType
ms.assetid: 8d331b09-745e-863d-a32c-77a9f1448b85
ms.date: 06/08/2017
---


# PictureType Property

Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart. For the Point and Series objects, read/write XlChartPictureType. For the LegendKey object, read/write Long. For the Floor and Walls objects, read/write Variant.



|XlChartPictureType can be one of these XlChartPictureType constants.|
| **xlStack**. Stacks the pictures to reach the necessary value.|
| **xlStretch**. Stretches the picture to reach the necessary value.|
| **xlStackScale**. Stacks the pictures; use the  **[PictureUnit](pictureunit-property.md)** property to determine what unit each picture represents.|

 _expression_. **PictureType**

 _expression_ Required. An expression that returns one of the above objects.

## Example

This example sets series one to stretch pictures. The example should be run on a 2-D column chart with picture data markers.


```
myChart.SeriesCollection(1).PictureType = xlStretch
```


