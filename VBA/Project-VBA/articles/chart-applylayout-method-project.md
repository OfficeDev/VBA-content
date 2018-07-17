---
title: Chart.ApplyLayout Method (Project)
keywords: vbapj.chm131609
f1_keywords:
- vbapj.chm131609
ms.prod: project-server
ms.assetid: 943ca7d6-aa2e-9058-f33b-4efd4138b497
ms.date: 06/08/2017
---


# Chart.ApplyLayout Method (Project)
Applies the specified chart layout and chart type to a selected chart.

## Syntax

 _expression_. **ApplyLayout** _(Layout,_ _varChartType)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Layout_|Required|**Long**|Specifies the type of layout, where the value corresponds to a  **Quick Layout** item on the ribbon.|
| _varChartType_|Optional|**Variant**|Can be one of the  **Office.XlChartType** constants.|
| _Layout_|Required|INT32||
| _varChartType_|Optional|VARIANT||

### Return value

 **Nothing**


## Remarks

When you select a chart in a report, the  **Quick Layout** drop-down list is on the **DESIGN** tab under **CHART TOOLS** on the ribbon. You can manually change the chart type by using the **Change Chart Type** command on the same ribbon tab.

When you use a layout on the current chart type, the  _Layout_ parameter is limited to the number of items in the **Quick Layout** drop-down list. You can use the _varChartType_ parameter to apply the layout of a different chart type on the current chart. For example, you can apply the layouts that are available from a line chart to a column chart. The **ApplyLayout** method adds only the line chart elements that are also available for the column chart type.


## Example

The following example applies the  **Quick Layout** item number 12 from a line chart to the selected chart.


```vb
Sub SetChartLayout()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.ApplyLayout Layout:=12, varChartType:=Office.XlChartType.xlLine
End Sub
```


## See also


#### Other resources


[Chart Object](chart-object-project.md)
