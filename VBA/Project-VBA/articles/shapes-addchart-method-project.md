---
title: Shapes.AddChart Method (Project)
ms.prod: project-server
ms.assetid: d404a9de-c1aa-c2a0-bf85-dc1f1735cf3c
ms.date: 06/08/2017
---


# Shapes.AddChart Method (Project)
Creates a chart at the specified location on the active report. Returns a  **Shape** object that represents the chart.

## Syntax

 _expression_. **AddChart** _(Style,_ _Type,_ _Left,_ _Top,_ _Width,_ _Height,_ _NewLayout)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Optional|**Integer**|Specifies the color style of the chart. The values correspond to the  **Change Colors** drop-down list on the **Chart Styles** group, on the **DESIGN** tab, under **CHART TOOLS** on the ribbon (but the values are not in the same order).|
| _Type_|Optional|**XlChartType**|The type of chart to add, such as a column chart or pie chart.|
| _Left_|Optional|**Single**|The position, measured in points, of the left edge of the chart.|
| _Top_|Optional|**Single**|The position, measured in points, of the top edge of the chart.|
| _Width_|Optional|**Single**|The width of the chart, measured in points.|
| _Height_|Optional|**Single**|The height of the chart, measured in points.|
| _NewLayout_|Optional|**Boolean**| _NewLayout_ is not used in Project.|
| _Style_|Optional|INT||
| _Type_|Optional|XLCHARTTYPE||
| _Left_|Optional|FLOAT||
| _Top_|Optional|FLOAT||
| _Width_|Optional|FLOAT||
| _Height_|Optional|FLOAT||
| _NewLayout_|Optional|BOOL||
|Name|Required/Optional|Data type|Description|

### Return value

 **Shape**


## Example

The following example creates a report that has a default bar chart type with orange-colored bars.


```vb
Sub AddDefaultChart()
    Dim chartReport As Report
    Dim reportName As String
    
    ' Add a report.
    reportName = "Test chart report"
    Set chartReport = ActiveProject.Reports.Add(reportName)

    ' Add a chart.
    Dim chartShape As shape
    Set chartShape = ActiveProject.Reports(reportName).Shapes.AddChart(Style:=12)
    
    With chartShape
        .Chart.SetElement msoElementChartTitleAboveChart
        .Chart.ChartTitle.Text = "Test Chart"
    End With
End Sub
```


## See also


#### Other resources


[Shapes Object](shapes-object-project.md)
[Shape Object](shape-object-project.md)
[Chart Object](chart-object-project.md)
