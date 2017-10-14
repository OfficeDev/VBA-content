---
title: Chart.ChartWizard Method (Project)
ms.prod: project-server
ms.assetid: 7626dc1f-89e1-3f18-0859-ebe2e0771de0
ms.date: 06/08/2017
---


# Chart.ChartWizard Method (Project)
Modifies the properties and formatting of a chart.

## Syntax

 _expression_. **ChartWizard** _(varSource,_ _varGallery,_ _varFormat,_ _varPlotBy,_ _varCategoryLabels,_ _varSeriesLabels,_ _varHasLegend,_ _varTitle,_ _varCategoryTitle,_ _varValueTitle,_ _varExtraTitle)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _varSource_|Optional|**Variant**|The source data for a new chart. If the  _varSource_ argument is omitted, Project edits the active report or the selected chart on the active report.|
| _varGallery_|Optional|**Variant**|One of the constants of the  **Office.XlChartType** enumeration, which specifies the chart type.|
| _varFormat_|Optional|**Variant**|The option number for the built-in autoformats. Can be a number from 1 through 10, depending on the gallery type. If the  _varFormat_ argument is omitted, Project chooses a default value based on the gallery type and data source.|
| _varPlotBy_|Optional|**Variant**|Specifies whether the data for each series is in rows or columns. Can be one of the following  **Office.XlRowCol** constants: **xlRows** or **xlColumns**.|
| _varCategoryLabels_|Optional|**Variant**|An integer that specifies the number of rows or columns within the source range that contain category labels. Values can be from 0 (zero) through one less than the maximum number of the corresponding categories or series.|
| _varSeriesLabels_|Optional|**Variant**|An integer that specifies the number of rows or columns within the source range that contain series labels. Values can be from 0 (zero) through one less than the maximum number of the corresponding categories or series.|
| _varHasLegend_|Optional|**Variant**|Set  **True** to include a legend.|
| _varTitle_|Optional|**Variant**|The chart title.|
| _varCategoryTitle_|Optional|**Variant**|The category axis title.|
| _varValueTitle_|Optional|**Variant**|The value axis title.|
| _varExtraTitle_|Optional|**Variant**|The series axis title for 3-D charts or the second value axis title for 2-D charts.|
| _varSource_|Optional|VARIANT||
| _varGallery_|Optional|VARIANT||
| _varFormat_|Optional|VARIANT||
| _varPlotBy_|Optional|VARIANT||
| _varCategoryLabels_|Optional|VARIANT||
| _varSeriesLabels_|Optional|VARIANT||
| _varHasLegend_|Optional|VARIANT||
| _varTitle_|Optional|VARIANT||
| _varCategoryTitle_|Optional|VARIANT||
| _varValueTitle_|Optional|VARIANT||
| _varExtraTitle_|Optional|VARIANT||

### Return value

 **Nothing**


## Remarks

You can use the  **ChartWizard** method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties. The[AutoFormat](chart-autoformat-method-project.md) method can do the same job as a call to **ChartWizard** that uses only the _varGallery_ and _varFormat_ parameters.

If the  _Source_ parameter is omitted and the selection isn't an embedded chart on the active report, or the active report does not contain a chart, the **ChartWizard** method fails and an error occurs.


## Example

The following example reformats the chart on the active report as a line chart, adds a legend, and adds category and value axis titles.


```vb
Sub TestChartWizard()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.ChartWizard varGallery:=xlLine, varHasLegend:=True, varCategoryTitle:="Task", varValueTitle:="Hours"
End Sub
```


## See also


#### Other resources


[Chart Object](chart-object-project.md)
[AutoFormat Method](chart-autoformat-method-project.md)
