---
title: Point Object (Excel)
keywords: vbaxl10.chm575072
f1_keywords:
- vbaxl10.chm575072
ms.prod: excel
api_name:
- Excel.Point
ms.assetid: 48ed9aec-2d29-ec4d-8e55-fca13982c358
ms.date: 06/08/2017
---


# Point Object (Excel)

Represents a single point in a series in a chart.


## Remarks

 The **Point** object is a member of the **[Points](points-object-excel.md)** collection. The **Points** collection contains all the points in one series.


## Example

Use  **[Points](series-points-method-excel.md)** ( _index_ ), where _index_ is the point index number, to return a single **Point** object. Points are numbered from left to right on the series. `Points(1)` is the leftmost point, and `Points(Points.Count)` is the rightmost point. The following example sets the marker style for the third point in series one in embedded chart one on worksheet one. The specified series must be a 2-D line, scatter, or radar series.


```
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
```


## Methods



|**Name**|
|:-----|
|[ApplyDataLabels](point-applydatalabels-method-excel.md)|
|[ClearFormats](point-clearformats-method-excel.md)|
|[Copy](point-copy-method-excel.md)|
|[Delete](point-delete-method-excel.md)|
|[Paste](point-paste-method-excel.md)|
|[PieSliceLocation](point-pieslicelocation-method-excel.md)|
|[Select](point-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](point-application-property-excel.md)|
|[ApplyPictToEnd](point-applypicttoend-property-excel.md)|
|[ApplyPictToFront](point-applypicttofront-property-excel.md)|
|[ApplyPictToSides](point-applypicttosides-property-excel.md)|
|[Creator](point-creator-property-excel.md)|
|[DataLabel](point-datalabel-property-excel.md)|
|[Explosion](point-explosion-property-excel.md)|
|[Format](point-format-property-excel.md)|
|[Has3DEffect](point-has3deffect-property-excel.md)|
|[HasDataLabel](point-hasdatalabel-property-excel.md)|
|[Height](point-height-property-excel.md)|
|[InvertIfNegative](point-invertifnegative-property-excel.md)|
|[Left](point-left-property-excel.md)|
|[MarkerBackgroundColor](point-markerbackgroundcolor-property-excel.md)|
|[MarkerBackgroundColorIndex](point-markerbackgroundcolorindex-property-excel.md)|
|[MarkerForegroundColor](point-markerforegroundcolor-property-excel.md)|
|[MarkerForegroundColorIndex](point-markerforegroundcolorindex-property-excel.md)|
|[MarkerSize](point-markersize-property-excel.md)|
|[MarkerStyle](point-markerstyle-property-excel.md)|
|[Name](point-name-property-excel.md)|
|[Parent](point-parent-property-excel.md)|
|[PictureType](point-picturetype-property-excel.md)|
|[PictureUnit2](point-pictureunit2-property-excel.md)|
|[SecondaryPlot](point-secondaryplot-property-excel.md)|
|[Shadow](point-shadow-property-excel.md)|
|[Top](point-top-property-excel.md)|
|[Width](point-width-property-excel.md)|
|[IsTotal](point-istotal-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
