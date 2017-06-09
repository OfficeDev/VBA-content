---
title: Chart.AutoFormat Method (Project)
ms.prod: project-server
ms.assetid: 1f560c0e-aed8-c989-9721-8e30595ae56e
ms.date: 06/08/2017
---


# Chart.AutoFormat Method (Project)
Changes the chart to a default format for another chart type.

## Syntax

 _expression_. **AutoFormat** _(rGallery,_ _varFormat)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _rGallery_|Required|**Long**|One of the constants of the  **Office.XlChartType** enumeration, which specifies the chart type.|
| _varFormat_|Optional|**Variant**|The option number for the built-in autoformats. Can be a number from 1 through 10, depending on the gallery type. If the  _varFormat_ argument is omitted, Project chooses a default value based on the gallery type and data source.|
| _rGallery_|Required|INT32||
| _varFormat_|Optional|VARIANT||

### Return value

 **Nothing**


## Remarks

The [ChartWizard](chart-chartwizard-method-project.md) method can do the same job as the **AutoFormat** method, although **ChartWizard** has more options.


## Example

The following example changes the chart to the default  **3-D Stacked Area** format.


```vb
Sub TestAutoFormat()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.AutoFormat Office.XlChartType.xl3DAreaStacked
End Sub
```


## See also


#### Other resources


[Chart Object](chart-object-project.md)
[ChartWizard Method](chart-chartwizard-method-project.md)
