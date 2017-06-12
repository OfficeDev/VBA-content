---
title: Chart.Perspective Property (Project)
ms.prod: project-server
ms.assetid: a6a07c7a-ca79-d6aa-e6ef-1aa26b716852
ms.date: 06/08/2017
---


# Chart.Perspective Property (Project)
Gets or sets a value that represents the perspective for the 3-D chart view. Read/write  **Long**.

## Syntax

 _expression_. **Perspective**

 _expression_ A variable that represents a **Chart** object.


## Remarks

The value of the  **Perspective** property must be between 0 and 100. **Perspective** is ignored if the[RightAngleAxes](chart-rightangleaxes-property-project.md) property is **True**.


## Example

The following example sets the perspective of the chart to 20. The example should be run on a 3-D chart.


```vb
Sub SetPerspective()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3-D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.RightAngleAxes = False
    chartShape.Chart.Perspective = 20
End Sub
```


## Property value

 **INT**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
