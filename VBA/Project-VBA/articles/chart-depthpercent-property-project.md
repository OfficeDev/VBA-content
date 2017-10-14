---
title: Chart.DepthPercent Property (Project)
ms.prod: project-server
ms.assetid: 868997e8-225c-5899-ccb0-71e1c8d9acfd
ms.date: 06/08/2017
---


# Chart.DepthPercent Property (Project)
Gets or sets the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write  **Long**.

## Syntax

 _expression_. **DepthPercent**

 _expression_ A variable that represents a **Chart** object.


## Remarks

The  **DepthPercent** property fails on 2-D charts.


## Example

The following example sets the depth of the specified chart to be 50 percent of its width. The example should be run on a 3-D chart.


```vb
Sub SetDepthPercent()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.DepthPercent = 50
End Sub
```


## Property value

 **INT**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
