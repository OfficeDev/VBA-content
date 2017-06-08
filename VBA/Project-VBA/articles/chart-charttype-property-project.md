---
title: Chart.ChartType Property (Project)
ms.prod: project-server
ms.assetid: c2557457-8aab-dec9-8098-e14b31a87c4f
ms.date: 06/08/2017
---


# Chart.ChartType Property (Project)
Gets or sets the chart type. Read/write  **Office.XlChartType**.

## Syntax

 _expression_. **ChartType**

 _expression_ A variable that represents a **Chart** object.


## Remarks

The  **ChartType** property corresponds to an action in the **Change Chart Type** dialog box. The command is on the ribbon under **CHART TOOLS**, on the  **DESIGN** tab.


## Example

The following example changes a clustered column chart to a clustered 3-D column chart type.


```vb
Sub SwitchChartTo3D()
    Dim chartShape As Shape
    
    Set chartShape = ActiveProject.Reports("Simple scalar chart").Shapes(1)
    
    With chartShape.Chart
        Debug.Print .ChartType
        
        If .ChartType = xlColumnClustered Then
            .ChartType = xl3DColumnClustered
        End If
    End With
End Sub
```


## Property value

 **XLCHARTTYPE**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
