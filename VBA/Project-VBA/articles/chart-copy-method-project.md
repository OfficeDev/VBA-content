---
title: Chart.Copy Method (Project)
keywords: vbapj.chm131611
f1_keywords:
- vbapj.chm131611
ms.prod: project-server
ms.assetid: 92627648-016a-0a69-52b8-bb24b1ea22d3
ms.date: 06/08/2017
---


# Chart.Copy Method (Project)
Copies a chart.

## Syntax

 _expression_. **Copy**

 _expression_ A variable that represents a **Chart** object.


### Return value

 **Variant**


## Example

The following example copies the chart and then pastes the chart as a picture on the active report.


```vb
Sub CopyAndPasteChart()
    Dim chartShape As Shape
    Dim reportName As String
    Dim duplicateChart As Chart
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.Copy
    Application.PasteAsPicture
End Sub
```


## See also


#### Other resources


[Chart Object](chart-object-project.md)
[CopyPicture Method](chart-copypicture-method-project.md)
