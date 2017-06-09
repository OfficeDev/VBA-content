---
title: Chart.Elevation Property (Project)
ms.prod: project-server
ms.assetid: c99cdc9b-3d3d-60c8-400f-6fa8818b4fd2
ms.date: 06/08/2017
---


# Chart.Elevation Property (Project)
Gets or sets the elevation of the 3-D chart view, in degrees. Read/write  **Long**.

## Syntax

 _expression_. **Elevation**

 _expression_ A variable that represents a **Chart** object.


## Remarks

The chart elevation is the angle from the horizontal at which you view the chart, in degrees. The default is 15 degrees for most chart types. The value of the  **Elevation** property must be between -90 and 90, except for 3-D bar charts, where it must be between 0 and 44. The **Elevation** property fails on 2-D charts.


## Example

The following example sets the elevation of the chart to 34 degrees. The example should be run on a 3-D chart.


```vb
Sub SetElevation()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.Elevation = 34
End Sub
```


## Property value

 **INT**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
