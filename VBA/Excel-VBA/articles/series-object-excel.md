---
title: Series Object (Excel)
keywords: vbaxl10.chm577072
f1_keywords:
- vbaxl10.chm577072
ms.prod: excel
api_name:
- Excel.Series
ms.assetid: c7d34b32-8172-f7a0-0a17-f01d44246b64
ms.date: 06/08/2017
---


# Series Object (Excel)

Represents a series in a chart.


## Remarks

 The **Series** object is a member of the **[SeriesCollection](seriescollection-object-excel.md)** collection.


## Example

Use  **SeriesCollection** ( _index_ ), where _index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series in embedded chart one on Sheet1.

The series index number indicates the order in which the series were added to the chart.  `SeriesCollection(1)` is the first series added to the chart, and `SeriesCollection(SeriesCollection.Count)` is the last one added.




```
Worksheets("sheet1").ChartObjects(1).Chart. _ 
 SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
```


## Methods



|**Name**|
|:-----|
|[ApplyDataLabels](series-applydatalabels-method-excel.md)|
|[ClearFormats](series-clearformats-method-excel.md)|
|[Copy](series-copy-method-excel.md)|
|[DataLabels](series-datalabels-method-excel.md)|
|[Delete](series-delete-method-excel.md)|
|[ErrorBar](series-errorbar-method-excel.md)|
|[Paste](series-paste-method-excel.md)|
|[Points](series-points-method-excel.md)|
|[Select](series-select-method-excel.md)|
|[Trendlines](series-trendlines-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](series-application-property-excel.md)|
|[ApplyPictToEnd](series-applypicttoend-property-excel.md)|
|[ApplyPictToFront](series-applypicttofront-property-excel.md)|
|[ApplyPictToSides](series-applypicttosides-property-excel.md)|
|[AxisGroup](series-axisgroup-property-excel.md)|
|[BarShape](series-barshape-property-excel.md)|
|[BubbleSizes](series-bubblesizes-property-excel.md)|
|[ChartType](series-charttype-property-excel.md)|
|[Creator](series-creator-property-excel.md)|
|[ErrorBars](series-errorbars-property-excel.md)|
|[Explosion](series-explosion-property-excel.md)|
|[Format](series-format-property-excel.md)|
|[Formula](series-formula-property-excel.md)|
|[FormulaLocal](series-formulalocal-property-excel.md)|
|[FormulaR1C1](series-formular1c1-property-excel.md)|
|[FormulaR1C1Local](series-formular1c1local-property-excel.md)|
|[Has3DEffect](series-has3deffect-property-excel.md)|
|[HasDataLabels](series-hasdatalabels-property-excel.md)|
|[HasErrorBars](series-haserrorbars-property-excel.md)|
|[HasLeaderLines](series-hasleaderlines-property-excel.md)|
|[InvertColor](series-invertcolor-property-excel.md)|
|[InvertColorIndex](series-invertcolorindex-property-excel.md)|
|[InvertIfNegative](series-invertifnegative-property-excel.md)|
|[IsFiltered](series-isfiltered-property-excel.md)|
|[LeaderLines](series-leaderlines-property-excel.md)|
|[MarkerBackgroundColor](series-markerbackgroundcolor-property-excel.md)|
|[MarkerBackgroundColorIndex](series-markerbackgroundcolorindex-property-excel.md)|
|[MarkerForegroundColor](series-markerforegroundcolor-property-excel.md)|
|[MarkerForegroundColorIndex](series-markerforegroundcolorindex-property-excel.md)|
|[MarkerSize](series-markersize-property-excel.md)|
|[MarkerStyle](series-markerstyle-property-excel.md)|
|[Name](series-name-property-excel.md)|
|[Parent](series-parent-property-excel.md)|
|[PictureType](series-picturetype-property-excel.md)|
|[PictureUnit2](series-pictureunit2-property-excel.md)|
|[PlotColorIndex](series-plotcolorindex-property-excel.md)|
|[PlotOrder](series-plotorder-property-excel.md)|
|[Shadow](series-shadow-property-excel.md)|
|[Smooth](series-smooth-property-excel.md)|
|[Type](series-type-property-excel.md)|
|[Values](series-values-property-excel.md)|
|[XValues](series-xvalues-property-excel.md)|
|[ParentDataLabelOption](series-parentdatalabeloption-property-excel.md)|
|[QuartileCalculationInclusiveMedian](series-quartilecalculationinclusivemedian-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
