---
title: Chart.ChartWizard Method (PowerPoint)
keywords: vbapp10.chm684020
f1_keywords:
- vbapp10.chm684020
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.ChartWizard
ms.assetid: aa13bff2-bead-2781-1f03-ea30ffe72a41
ms.date: 06/08/2017
---


# Chart.ChartWizard Method (PowerPoint)

Modifies the properties of the given chart. You can use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.

## Syntax

_expression_. **ChartWizard** (**_Source_**, **_Gallery_**, **_Format_**, **_PlotBy_**, **_CategoryLabels_**, **_SeriesLabels_**, **_HasLegend_**, **_Title_**, **_CategoryTitle_**, **_ValueTitle_**, **_ExtraTitle_**)

_expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Optional|**Variant**|The range that contains the source data for the new chart. If this argument is omitted, Word edits the active chart sheet or the selected chart on the active worksheet.|
| _Gallery_|Optional|**Variant**|One of the **[XlChartType](http://msdn.microsoft.com/library/bba4ee89-ee91-f55a-d2e0-59a73e5bfabe%28Office.15%29.aspx)** constants that specifies the chart type.|
| _Format_|Optional|**Variant**|The option number for the built-in autoformats. Can be a number from 1 through 10, depending on the gallery type. If this argument is omitted, Word chooses a default value based on the gallery type and data source.|
| _PlotBy_|Optional|**Variant**|Specifies whether the data for each series is in rows or columns. Can be one of the following **[XlRowCol](xlrowcol-enumeration-powerpoint.md)** constants: **xlRows** or **xlColumns**.|
| _CategoryLabels_|Optional|**Variant**|An integer that specifies the number of rows or columns within the source range that contain category labels. Allowed values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.|
| _SeriesLabels_|Optional|**Variant**|An integer that specifies the number of rows or columns within the source range that contain series labels. Allowed values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.|
| _HasLegend_|Optional|**Variant**|**True** to include a legend.|
| _Title_|Optional|**Variant**|The chart title text.|
| _CategoryTitle_|Optional|**Variant**|The category axis title text.|
| _ValueTitle_|Optional|**Variant**|The value axis title text.|
| _ExtraTitle_|Optional|**Variant**| The series axis title for 3-D charts or the second value axis title for 2-D charts.|

<br/>

## Remarks

If the Source parameter is omitted, and the selection is not a chart on the active document, this method fails and an error occurs.


## Example

**Note**  Although the following code applies to Word, you can readily modify it to apply to PowerPoint.

The following example reformats the first chart as a line chart, adds a legend, and adds category and value axis titles.

```vb
With ActiveDocument.InlineShapes(1).Chart 
    .ChartWizard _ 
        Gallery:=xlLine, _ 
        HasLegend:=True, _ 
        CategoryTitle:="Year", _ 
        ValueTitle:="Sales" 
End With
```


## See also

#### Concepts

- [Chart Object](chart-object-powerpoint.md)

