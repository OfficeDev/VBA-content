---
title: Chart.AutoScaling Property (Project)
ms.prod: project-server
ms.assetid: d7e1c8f7-8a2b-0474-1b4a-28a63605e929
ms.date: 06/08/2017
---


# Chart.AutoScaling Property (Project)
 **True** if Project scales a 3-D chart so that it is closer in size to the equivalent 2-D chart. Read/write **Boolean**.

## Syntax

 _expression_. **AutoScaling**

 _expression_ A variable that represents a **Chart** object.


## Remarks

For auto-scaling to work, the  **[RightAngleAxes](chart-rightangleaxes-property-project.md)** property must also be **True**. 


## Example

In the following example, the chart is the first shape in the "3-D chart" report. The example automatically scales the chart. The example should be run on a 3-D chart.


```vb
Sub SetChartColor()
    Dim chartShape As Shape
    
    Set chartShape = ActiveProject.Reports("3-D chart").Shapes(1)
    With chartShape
        .RightAngleAxes = True
        .AutoScaling = True
    End With
End Sub
```


## Property value

 **BOOL**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
[RightAngleAxes Property](chart-rightangleaxes-property-project.md)
