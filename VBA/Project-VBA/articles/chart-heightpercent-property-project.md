---
title: Chart.HeightPercent Property (Project)
ms.prod: project-server
ms.assetid: cb7e3a55-eb99-b02d-2242-ebdcbd954b35
ms.date: 06/08/2017
---


# Chart.HeightPercent Property (Project)
Gets or sets the height of a 3-D chart as a percentage of the chart width. Read/write  **Long**.

## Syntax

 _expression_. **HeightPercent**

 _expression_ A variable that represents a **Chart** object.


## Remarks

The  **HeightPercent** value can be between 5 and 500 percent.


## Example

The following example sets the height of the chart to 80 percent of its width. The example should be run on a 3-D chart.


```vb
Sub SetHeightPercent()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.HeightPercent = 80
End Sub
```


## Property value

 **INT**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
