---
title: Chart.ChartGroups Property (Project)
keywords: vbapj.chm131625
f1_keywords:
- vbapj.chm131625
ms.prod: project-server
ms.assetid: 49e50578-3b97-4bc5-6037-3d32f0f321a7
ms.date: 06/08/2017
---


# Chart.ChartGroups Property (Project)
Gets an object that represents either a single chart group or a collection of chart groups, where a chart group represents one or more series of data points that are plotted with the same format. Read-only  **Object**.

## Syntax

 _expression_. **ChartGroups**

 _expression_ A variable that represents a **Chart** object.


## Remarks

A chart contains one or more chart groups, and each chart group contains one or more series of data points. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format.


## Example

The following example should be run on a simple line chart. The example toggles drop lines on and off for the chart.


```vb
Sub ToggleDropLines()
    Dim chartShape As Shape
    Dim chartGroup As Office.IMsoChartGroup
    Dim dropLines As Boolean
    Dim reportName As String
    
    reportName = "Simple line chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    Set chartGroup = chartShape.Chart.ChartGroups(1)
    dropLines = chartGroup.HasDropLines
    
    MsgBox "Chart group in " &; reportName &; ": " _
        &; vbCrLf &; "Drop lines: " &; dropLines
        
    chartGroup.HasDropLines = Not dropLines
End Sub
```


## Property value

 **OBJECT**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
