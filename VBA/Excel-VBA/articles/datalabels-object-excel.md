---
title: DataLabels Object (Excel)
keywords: vbaxl10.chm583072
f1_keywords:
- vbaxl10.chm583072
ms.prod: excel
api_name:
- Excel.DataLabels
ms.assetid: 3d79271e-c702-e785-6984-d838d060a8c5
ms.date: 06/08/2017
---


# DataLabels Object (Excel)

A collection of all the  **[DataLabel](datalabel-object-excel.md)** objects for the specified series.


## Remarks

 Each **DataLabel** object represents a data label for a point or trendline. For a series without definable points (such as an area series), the **DataLabels** collection contains a single data label.


## Example

Use the  **[DataLabels](series-datalabels-method-excel.md)** method to return the **DataLabels** collection. The following example sets the number format for data labels on series one on chart sheet one.


```
With Charts(1).SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.NumberFormat = "##.##" 
End With
```

Use  **DataLabels** ( _index_ ), where _index_ is the data-label index number, to return a single **DataLabel** object. The following example sets the number format for the fifth data label in series one in embedded chart one on worksheet one.




```
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
```


## Methods



|**Name**|
|:-----|
|[Delete](datalabels-delete-method-excel.md)|
|[Item](datalabels-item-method-excel.md)|
|[Propagate](datalabels-propagate-method-excel.md)|
|[Select](datalabels-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](datalabels-application-property-excel.md)|
|[AutoText](datalabels-autotext-property-excel.md)|
|[Count](datalabels-count-property-excel.md)|
|[Creator](datalabels-creator-property-excel.md)|
|[Format](datalabels-format-property-excel.md)|
|[HorizontalAlignment](datalabels-horizontalalignment-property-excel.md)|
|[Name](datalabels-name-property-excel.md)|
|[NumberFormat](datalabels-numberformat-property-excel.md)|
|[NumberFormatLinked](datalabels-numberformatlinked-property-excel.md)|
|[NumberFormatLocal](datalabels-numberformatlocal-property-excel.md)|
|[Orientation](datalabels-orientation-property-excel.md)|
|[Parent](datalabels-parent-property-excel.md)|
|[Position](datalabels-position-property-excel.md)|
|[ReadingOrder](datalabels-readingorder-property-excel.md)|
|[Separator](datalabels-separator-property-excel.md)|
|[Shadow](datalabels-shadow-property-excel.md)|
|[ShowBubbleSize](datalabels-showbubblesize-property-excel.md)|
|[ShowCategoryName](datalabels-showcategoryname-property-excel.md)|
|[ShowLegendKey](datalabels-showlegendkey-property-excel.md)|
|[ShowPercentage](datalabels-showpercentage-property-excel.md)|
|[ShowRange](datalabels-showrange-property-excel.md)|
|[ShowSeriesName](datalabels-showseriesname-property-excel.md)|
|[ShowValue](datalabels-showvalue-property-excel.md)|
|[VerticalAlignment](datalabels-verticalalignment-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
