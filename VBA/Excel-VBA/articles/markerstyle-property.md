---
title: MarkerStyle Property
keywords: vbagr10.chm65608
f1_keywords:
- vbagr10.chm65608
ms.prod: excel
api_name:
- Excel.MarkerStyle
ms.assetid: 6010628c-55ab-a613-efb0-53e6abb92295
ms.date: 06/08/2017
---


# MarkerStyle Property

Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle .



|XlMarkerStyle can be one of these XlMarkerStyle constants.|
| **xlMarkerStyleCircle**. Circular markers|
| **xlMarkerStyleDiamond**. Diamond-shaped markers |
| **xlMarkerStyleNone**. No markers|
| **xlMarkerStylePlus**. Square markers with a plus sign|
| **xlMarkerStyleStar**. Square markers with an asterisk|
| **xlMarkerStyleX**. Square markers with an X|
| **xlMarkerStyleAutomatic**. Automatic markers|
| **xlMarkerStyleDash**. Long bar markers|
| **xlMarkerStyleDot**. Short bar markers|
| **xlMarkerStylePicture**. Picture markers|
| **xlMarkerStyleSquare**. Square markers|
| **xlMarkerStyleTriangle**. Triangular markers|

 _expression_. **MarkerStyle**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the marker style for series one. The example should be run on a 2-D line chart.


```
myChart.SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
```


