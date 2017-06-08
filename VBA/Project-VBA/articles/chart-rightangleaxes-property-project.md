---
title: Chart.RightAngleAxes Property (Project)
ms.prod: project-server
ms.assetid: 51e8cde1-53c7-90ff-b5c7-72a091461f6b
ms.date: 06/08/2017
---


# Chart.RightAngleAxes Property (Project)
 **True** if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3-D line, 3-D column, and 3-D bar charts. Read/write **Boolean**.

## Syntax

 _expression_. **RightAngleAxes**

 _expression_ A variable that represents a **Chart** object.


## Remarks

If the  **RightAngleAxes** property is **True**, the  **[Perspective](chart-perspective-property-project.md)** property is ignored.


## Example

The following example sets the chart axes to intersect at right angles. The example should be run on a 3-D chart.


```vb
Sub SetRightAngleAxes()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3-D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.RightAngleAxes = True
End Sub
```


## Property value

 **VARIANT**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
[AutoScaling Property](chart-autoscaling-property-project.md)
