---
title: ApplyCustomType Method
keywords: vbagr10.chm66937
f1_keywords:
- vbagr10.chm66937
ms.prod: excel
api_name:
- Excel.ApplyCustomType
ms.assetid: 5385d195-96ce-bdd3-e84d-596fd4236904
ms.date: 06/08/2017
---


# ApplyCustomType Method

ApplyCustomType method as it applies to the  **Series** object.

Applies a standard or custom chart type to a series.

 _expression_. **ApplyCustomType**( **_ChartType_**)

 _expression_ Required. An expression that returns one of the above objects.
 **ChartType**Required 
 **XlChartType**
. A standard chart type.


|XlChartType can be one of these XlChartType constants.|
| **xl3DAreaStacked**|
| **xlConeCol**|
| **xlConeColStacked**|
| **xlCylinderBarClustered**|
| **xlCylinderBarStacked100**|
| **xlCylinderColClustered**|
| **xlCylinderColStacked100**|
| **xlDoughnutExploded**|
| **xlLineMarkers**|
| **xlLineMarkersStacked100**|
| **xlLineStacked100**|
| **xlPieExploded**|
| **xlPyramidBarClustered**|
| **xlPyramidBarStacked100**|
| **xlPyramidColClustered**|
| **xlPyramidColStacked100**|
| **xlRadarFilled**|
| **xlStockHLC**|
| **xlStockVHLC**|
| **xlSurface**|
| **xlSurfaceTopViewWireframe**|
| **xlXYScatter**|
| **xlXYScatterLinesNoMarkers**|
| **xlXYScatterSmoothNoMarkers**|
| **xl3DArea**|
| **xl3DAreaStacked100**|
| **xl3DBarClustered**|
| **xl3DBarStacked**|
| **xl3DBarStacked100**|
| **xl3DColumn**|
| **xl3DColumnClustered**|
| **xl3DColumnStacked**|
| **xl3DColumnStacked100**|
| **xl3DLine**|
| **xl3DPie**|
| **xl3DPieExploded**|
| **xlArea**|
| **xlAreaStacked**|
| **xlAreaStacked100**|
| **xlBarClustered**|
| **xlBarOfPie**|
| **xlBarStacked**|
| **xlBarStacked100**|
| **xlBubble**|
| **xlBubble3DEffect**|
| **xlColumnClustered**|
| **xlColumnStacked**|
| **xlColumnStacked100**|
| **xlConeBarClustered**|
| **xlConeBarStacked**|
| **xlConeBarStacked100**|
| **xlConeColClustered**|
| **xlConeColStacked100**|
| **xlCylinderBarStacked**|
| **xlCylinderCol**|
| **xlCylinderColStacked**|
| **xlDoughnut**|
| **xlLine**|
| **xlLineMarkersStacked**|
| **xlLineStacked**|
| **xlPie**|
| **xlPieOfPie**|
| **xlPyramidBarStacked**|
| **xlPyramidCol**|
| **xlPyramidColStacked**|
| **xlRadar**|
| **xlRadarMarkers**|
| **xlStockOHLC**|
| **xlStockVOHLC**|
| **xlSurfaceTopView**|
| **xlSurfaceWireframe**|
| **xlXYScatterLines**|
| **xlXYScatterSmooth**|
ApplyCustomType method as it applies to the  **Chart** object.
Applies a standard or custom chart type to a chart.
 _expression_. **ApplyCustomType**( **_ChartType_**,  **_TypeName_**)
 _expression_ Required. An expression that returns one of the above objects.
 **ChartType**Required 
 **XlChartType**
. A standard chart type.


|XlChartType can be one of these XlChartType constants.|
| **xl3DAreaStacked**|
| **xlConeCol**|
| **xlConeColStacked**|
| **xlCylinderBarClustered**|
| **xlCylinderBarStacked100**|
| **xlCylinderColClustered**|
| **xlCylinderColStacked100**|
| **xlDoughnutExploded**|
| **xlLineMarkers**|
| **xlLineMarkersStacked100**|
| **xlLineStacked100**|
| **xlPieExploded**|
| **xlPyramidBarClustered**|
| **xlPyramidBarStacked100**|
| **xlPyramidColClustered**|
| **xlPyramidColStacked100**|
| **xlRadarFilled**|
| **xlStockHLC**|
| **xlStockVHLC**|
| **xlSurface**|
| **xlSurfaceTopViewWireframe**|
| **xlXYScatter**|
| **xlXYScatterLinesNoMarkers**|
| **xlXYScatterSmoothNoMarkers**|
| **xl3DArea**|
| **xl3DAreaStacked100**|
| **xl3DBarClustered**|
| **xl3DBarStacked**|
| **xl3DBarStacked100**|
| **xl3DColumn**|
| **xl3DColumnClustered**|
| **xl3DColumnStacked**|
| **xl3DColumnStacked100**|
| **xl3DLine**|
| **xl3DPie**|
| **xl3DPieExploded**|
| **xlArea**|
| **xlAreaStacked**|
| **xlAreaStacked100**|
| **xlBarClustered**|
| **xlBarOfPie**|
| **xlBarStacked**|
| **xlBarStacked100**|
| **xlBubble**|
| **xlBubble3DEffect**|
| **xlColumnClustered**|
| **xlColumnStacked**|
| **xlColumnStacked100**|
| **xlConeBarClustered**|
| **xlConeBarStacked**|
| **xlConeBarStacked100**|
| **xlConeColClustered**|
| **xlConeColStacked100**|
| **xlCylinderBarStacked**|
| **xlCylinderCol**|
| **xlCylinderColStacked**|
| **xlDoughnut**|
| **xlLine**|
| **xlLineMarkersStacked**|
| **xlLineStacked**|
| **xlPie**|
| **xlPieOfPie**|
| **xlPyramidBarStacked**|
| **xlPyramidCol**|
| **xlPyramidColStacked**|
| **xlRadar**|
| **xlRadarMarkers**|
| **xlStockOHLC**|
| **xlStockVOHLC**|
| **xlSurfaceTopView**|
| **xlSurfaceWireframe**|
| **xlXYScatterLines**|
| **xlXYScatterSmooth**|
 **TypeName**Optional  **Variant**. A  **String** naming the custom chart type when **_ChartType_** specifies a custom chart gallery.

## Example

This example applies the line with the markers chart type.


```
myChart.ApplyCustomType xlLineMarkers
```


