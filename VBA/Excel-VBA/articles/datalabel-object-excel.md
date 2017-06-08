---
title: DataLabel Object (Excel)
keywords: vbaxl10.chm581072
f1_keywords:
- vbaxl10.chm581072
ms.prod: excel
api_name:
- Excel.DataLabel
ms.assetid: bb342572-8761-b326-548a-98455172f9a8
ms.date: 06/08/2017
---


# DataLabel Object (Excel)

Represents the data label on a chart point or trendline.


## Remarks

 On a series, the **DataLabel** object is a member of the **[DataLabels](datalabels-object-excel.md)** collection. The **DataLabels** collection contains a **DataLabel** object for each point. For a series without definable points (such as an area series), the **DataLabels** collection contains a single **DataLabel** object.


## Example

Use  **[DataLabels](series-datalabels-method-excel.md)** ( _index_ ), where _index_ is the data-label index number, to return a single **DataLabel** object. The following example sets the number format for the fifth data label in series one in embedded chart one on worksheet one.


```
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
```

Use the  **[DataLabel](point-datalabel-property-excel.md)** property to return the **DataLabel** object for a single point. The following example turns on the data label for the second point in series one on the chart sheet named "Chart1" and sets the data label text to "Saturday."




```
With Charts("chart1") 
 With .SeriesCollection(1).Points(2) 
 .HasDataLabel = True 
 .DataLabel.Text = "Saturday" 
 End With 
End With
```

On a trendline, the  **[DataLabel](trendline-datalabel-property-excel.md)** property returns the text shown with the trendline. This can be the equation, the R-squared value, or both (if both are showing). The following example sets the trendline text to show only the equation and then places the data label text in cell A1 on the worksheet named "Sheet1."




```
With Charts("chart1").SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = False 
 .DisplayEquation = True 
 Worksheets("sheet1").Range("a1").Value = .DataLabel.Text 
End With
```


## Methods



|**Name**|
|:-----|
|[Delete](datalabel-delete-method-excel.md)|
|[Select](datalabel-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](datalabel-application-property-excel.md)|
|[AutoText](datalabel-autotext-property-excel.md)|
|[Caption](datalabel-caption-property-excel.md)|
|[Characters](datalabel-characters-property-excel.md)|
|[Creator](datalabel-creator-property-excel.md)|
|[Format](datalabel-format-property-excel.md)|
|[Formula](datalabel-formula-property-excel.md)|
|[FormulaLocal](datalabel-formulalocal-property-excel.md)|
|[FormulaR1C1](datalabel-formular1c1-property-excel.md)|
|[FormulaR1C1Local](datalabel-formular1c1local-property-excel.md)|
|[Height](datalabel-height-property-excel.md)|
|[HorizontalAlignment](datalabel-horizontalalignment-property-excel.md)|
|[Left](datalabel-left-property-excel.md)|
|[Name](datalabel-name-property-excel.md)|
|[NumberFormat](datalabel-numberformat-property-excel.md)|
|[NumberFormatLinked](datalabel-numberformatlinked-property-excel.md)|
|[NumberFormatLocal](datalabel-numberformatlocal-property-excel.md)|
|[Orientation](datalabel-orientation-property-excel.md)|
|[Parent](datalabel-parent-property-excel.md)|
|[Position](datalabel-position-property-excel.md)|
|[ReadingOrder](datalabel-readingorder-property-excel.md)|
|[Separator](datalabel-separator-property-excel.md)|
|[Shadow](datalabel-shadow-property-excel.md)|
|[ShowBubbleSize](datalabel-showbubblesize-property-excel.md)|
|[ShowCategoryName](datalabel-showcategoryname-property-excel.md)|
|[ShowLegendKey](datalabel-showlegendkey-property-excel.md)|
|[ShowPercentage](datalabel-showpercentage-property-excel.md)|
|[ShowRange](datalabel-showrange-property-excel.md)|
|[ShowSeriesName](datalabel-showseriesname-property-excel.md)|
|[ShowValue](datalabel-showvalue-property-excel.md)|
|[Text](datalabel-text-property-excel.md)|
|[Top](datalabel-top-property-excel.md)|
|[VerticalAlignment](datalabel-verticalalignment-property-excel.md)|
|[Width](datalabel-width-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
