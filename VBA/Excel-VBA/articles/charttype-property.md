---
title: ChartType Property
keywords: vbagr10.chm66936
f1_keywords:
- vbagr10.chm66936
ms.prod: excel
api_name:
- Excel.ChartType
ms.assetid: a59871a9-d2f9-657a-1553-eba8c4e4a5a8
ms.date: 06/08/2017
---


# ChartType Property

Returns or sets the chart type. Read/write XlChartType .



|XlChartType can be one of these XlChartType constants.|
| **xl3DArea**. 3-D Area|
| **xl3DAreaStacked**. 3-D Stacked Area|
| **xl3DAreaStacked100**. 3-D Stacked Area|
| **xl3DBarClustered**. 3-D Clustered Bar|
| **xl3DBarStacked**. 3-D Stacked Bar|
| **xl3DBarStacked100**. 3-D 100% Stacked Bar|
| **xl3DColumn**. 3-D Column|
| **xl3DColumnClustered**. 3-D Clustered Column|
| **xl3DColumnStacked**. 3-D Stacked Column|
| **xl3DColumnStacked100**. 3-D 100% Stacked Column|
| **xl3DLine**. 3-D Line|
| **xl3DPie**. 3-D Pie|
| **xl3DPieExploded**. Exploded 3-D Pie|
| **xlArea**. Area|
| **xlAreaStacked**. Stacked Area|
| **xlAreaStacked100**. 100% Stacked Area|
| **xlBarClustered**. Clustered Bar|
| **xlBarOfPie**. Bar of Pie|
| **xlBarStacked**. Stacked Bar|
| **xlBarStacked100**. 100% Stacked Bar|
| **xlBubble**. Bubble|
| **xlBubble3DEffect**. Bubble with 3-D Effects|
| **xlColumnClustered**. Clustered Column|
| **xlColumnStacked**. Stacked Column|
| **xlColumnStacked100**. 100% Stacked Column|
| **xlConeBarClustered**. Clustered Cone Bar|
| **xlConeBarStacked**. Stacked Cone Bar|
| **xlConeBarStacked100**. 100% Stacked Cone Bar|
| **xlConeCol**. 3-D Cone Column|
| **xlConeColClustered**. Clustered Cone Column|
| **xlConeColStacked**. Stacked Cone Column|
| **xlConeColStacked100**. 100% Stacked Cone Column|
| **xlCylinderBarStacked**. Stacked Cylinder Bar|
| **xlCylinderCol**. 3-D Cylinder Column|
| **xlCylinderColStacked**. Stacked Cylinder Column|
| **xlCylinderBarClustered**. Clustered Cylinder Bar|
| **xlCylinderBarStacked100**. 100% Stacked Cylinder Bar|
| **xlCylinderColClustered**. Clustered Cylinder Column|
| **xlCylinderColStacked100**. 100% Stacked Cylinder Column|
| **xlDoughnut**. Doughnut|
| **xlDoughnutExploded**. Exploded Doughnut|
| **xlLineMarkers**. Line with Data Markers|
| **xlLineMarkersStacked100**. 100% Stacked Line with Markers|
| **xlLineStacked100**. 100% Stacked Line|
| **xlLine**. Line|
| **xlLineMarkersStacked**. Stacked Line with Data Markers|
| **xlLineStacked**. Stacked Line|
| **xlPie**. Pie|
| **xlPieExploded**. Exploded Pie|
| **xlPieOfPie**. Pie of Pie|
| **xlPyramidBarClustered**. Clustered Pyramid Bar|
| **xlPyramidBarStacked**. Stacked Pyramid Bar|
| **xlPyramidBarStacked100**. 100% Stacked Pyramid Bar|
| **xlPyramidCol**. 3-D Pyramid Column|
| **xlPyramidColStacked**. Stacked Pyramid Column|
| **xlPyramidColClustered**. Clustered Pyramid Column|
| **xlPyramidColStacked100**. 100% Stacked Pyramid Column|
| **xlRadar**. Radar|
| **xlRadarFilled**. Filled Radar|
| **xlRadarMarkers**. Radar with Data Markers|
| **xlStockHLC**. High-Low-Close|
| **xlStockOHLC**. Open-High-Low-Close|
| **xlStockVHLC**. Volume-High-Low-Close|
| **xlStockVOHLC**. Volume-Open-High-Low-Close|
| **xlSurface**. 3-D Surface|
| **xlSurfaceTopView**. Surface (Top View)|
| **xlSurfaceTopViewWireframe**. Surface (Top View wire-frame)|
| **xlSurfaceWireframe**. 3-D Surface(wire-frame)|
| **xlXYScatter**. Scatter|
| **xlXYScatterLines**. Scatter with Lines|
| **xlXYScatterLinesNoMarkers**. Scatter with Lines and No Data Markers|
| **xlXYScatterSmooth**. Scatter with SmoothedLines|
| **xlXYScatterSmoothNoMarkers**. Scatter with Smoothed Lines and No Data Markers|

 _expression_. **ChartType**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the bubble size in chart group one to 200 percent of the default size if the chart is a 2-D bubble chart.


```vb
With myChart 
 If .ChartType = xlBubble Then 
 .ChartGroups(1).BubbleScale = 200 
 End If 
End With
```


