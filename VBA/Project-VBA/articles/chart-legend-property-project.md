---
title: Chart.Legend Property (Project)
keywords: vbapj.chm131629
f1_keywords:
- vbapj.chm131629
ms.prod: project-server
ms.assetid: 38c3332c-6087-4f7b-5c02-31cba5c6933f
ms.date: 06/08/2017
---


# Chart.Legend Property (Project)
Gets an  **Office.IMsoLegend** object that represents the legend for a chart. Read-only **IMsoLegend**.

## Syntax

 _expression_. **Legend**

 _expression_ A variable that represents a **Chart** object.


## Example

The following example turns on the legend for the chart, and then sets the top of the legend 20 points lower than it was set previously.


```vb
Sub SetLegendTop()
    Dim chartShape As Shape
    Dim reportName As String
    Dim legendTop As Double
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.HasLegend = True
    legendTop = chartShape.Chart.Legend.Top
    chartShape.Chart.Legend.Top = legendTop + 20
End Sub
```


## Property value

 **IMSOLEGEND**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
