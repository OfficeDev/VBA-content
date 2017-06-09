---
title: Chart Object (Word)
keywords: vbawd10.chm1211
f1_keywords:
- vbawd10.chm1211
ms.prod: word
api_name:
- Word.Chart
ms.assetid: 366a825e-0daf-dbb7-b6f2-e7ce1a5ee2ef
ms.date: 06/08/2017
---


# Chart Object (Word)

Represents a chart in a document.


## Remarks

The Example section describes the following properties and methods for returning a  **Chart** object:




- The  **[Chart](inlineshape-chart-property-word.md)** property.
    
- The  **[AddChart](http://msdn.microsoft.com/library/1b168e7b-543a-a817-51b0-8171beecc946%28Office.15%29.aspx)** method.
    



## Example

The  **[InlineShapes](inlineshapes-object-word.md)** collection contains an object for each inline shape, including charts, in a document. Use **InlineShapes** ( _Index_ ), where Index is the index number of an inline shape, to return a single **InlineShape** object. Use the **[HasChart](inlineshape-haschart-property-word.md)** property to determine whether the **InlineShape** object represents a chart. If the **HasChart** property is set to **True**, use the **[Chart](inlineshape-chart-property-word.md)** property to return a **Chart** object.

You can also use the  **[Type](inlineshape-type-property-word.md)** property to determine whether the **InlineShape** object represents a chart. If the **Type** property is set to **WdInlineShapeChart**, the inline shape represents a chart.

The following example verifies whether the first inline shape in the active document represents a chart. If so, the example changes the fore color of the first series on the chart.




```
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed 
 End If 
End With
```

The following example creates a new 3-D column chart and adds it to the active document.




```
ActiveDocument.InlineShapes.AddChart Type:=xl3DColumn 

```


## Methods



|**Name**|
|:-----|
|[ApplyChartTemplate](chart-applycharttemplate-method-word.md)|
|[ApplyDataLabels](chart-applydatalabels-method-word.md)|
|[ApplyLayout](chart-applylayout-method-word.md)|
|[Axes](chart-axes-method-word.md)|
|[ChartWizard](chart-chartwizard-method-word.md)|
|[ClearToMatchColorStyle](chart-cleartomatchcolorstyle-method-word.md)|
|[ClearToMatchStyle](chart-cleartomatchstyle-method-word.md)|
|[Copy](chart-copy-method-word.md)|
|[CopyPicture](chart-copypicture-method-word.md)|
|[Delete](chart-delete-method-word.md)|
|[Export](chart-export-method-word.md)|
|[FullSeriesCollection](chart-fullseriescollection-method-word.md)|
|[GetChartElement](chart-getchartelement-method-word.md)|
|[Paste](chart-paste-method-word.md)|
|[Refresh](chart-refresh-method-word.md)|
|[SaveChartTemplate](chart-savecharttemplate-method-word.md)|
|[Select](chart-select-method-word.md)|
|[SeriesCollection](chart-seriescollection-method-word.md)|
|[SetBackgroundPicture](chart-setbackgroundpicture-method-word.md)|
|[SetDefaultChart](chart-setdefaultchart-method-word.md)|
|[SetElement](chart-setelement-method-word.md)|
|[SetSourceData](chart-setsourcedata-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](chart-application-property-word.md)|
|[AutoScaling](chart-autoscaling-property-word.md)|
|[BackWall](chart-backwall-property-word.md)|
|[BarShape](chart-barshape-property-word.md)|
|[CategoryLabelLevel](chart-categorylabellevel-property-word.md)|
|[ChartArea](chart-chartarea-property-word.md)|
|[ChartColor](chart-chartcolor-property-word.md)|
|[ChartData](chart-chartdata-property-word.md)|
|[ChartGroups](chart-chartgroups-property-word.md)|
|[ChartStyle](chart-chartstyle-property-word.md)|
|[ChartTitle](chart-charttitle-property-word.md)|
|[ChartType](chart-charttype-property-word.md)|
|[Creator](chart-creator-property-word.md)|
|[DataTable](chart-datatable-property-word.md)|
|[DepthPercent](chart-depthpercent-property-word.md)|
|[DisplayBlanksAs](chart-displayblanksas-property-word.md)|
|[Elevation](chart-elevation-property-word.md)|
|[Floor](chart-floor-property-word.md)|
|[GapDepth](chart-gapdepth-property-word.md)|
|[HasAxis](chart-hasaxis-property-word.md)|
|[HasDataTable](chart-hasdatatable-property-word.md)|
|[HasLegend](chart-haslegend-property-word.md)|
|[HasTitle](chart-hastitle-property-word.md)|
|[HeightPercent](chart-heightpercent-property-word.md)|
|[Legend](chart-legend-property-word.md)|
|[Parent](chart-parent-property-word.md)|
|[Perspective](chart-perspective-property-word.md)|
|[PivotLayout](chart-pivotlayout-property-word.md)|
|[PlotArea](chart-plotarea-property-word.md)|
|[PlotBy](chart-plotby-property-word.md)|
|[PlotVisibleOnly](chart-plotvisibleonly-property-word.md)|
|[RightAngleAxes](chart-rightangleaxes-property-word.md)|
|[Rotation](chart-rotation-property-word.md)|
|[SeriesNameLevel](chart-seriesnamelevel-property-word.md)|
|[Shapes](chart-shapes-property-word.md)|
|[ShowAllFieldButtons](chart-showallfieldbuttons-property-word.md)|
|[ShowAxisFieldButtons](chart-showaxisfieldbuttons-property-word.md)|
|[ShowDataLabelsOverMaximum](chart-showdatalabelsovermaximum-property-word.md)|
|[ShowLegendFieldButtons](chart-showlegendfieldbuttons-property-word.md)|
|[ShowReportFilterFieldButtons](chart-showreportfilterfieldbuttons-property-word.md)|
|[ShowValueFieldButtons](chart-showvaluefieldbuttons-property-word.md)|
|[SideWall](chart-sidewall-property-word.md)|
|[Walls](chart-walls-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
