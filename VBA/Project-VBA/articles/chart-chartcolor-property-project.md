---
title: Chart.ChartColor Property (Project)
ms.prod: project-server
ms.assetid: bd8b5b9c-abfe-761b-a4bd-2978c43b9565
ms.date: 06/08/2017
---


# Chart.ChartColor Property (Project)
Gets or sets the index of chart colors for the active chart. Read/write  **Variant**.

## Syntax

 _expression_. **ChartColor**

 _expression_ A variable that represents a **Chart** object.


## Remarks

The  **ChartColor** property corresponds to a selection in the **Change Colors** drop-down list, which is on the ribbon under **Chart Tools**, on the  **Format** tab, in the **ChartStyles** group.


 **Note**  The  **Colors** drop-down list on the ribbon under **REPORT TOOLS**, on the  **DESIGN** tab, in the **Themes** group, changes the color theme of the entire report, including any charts on the report. The VBA object model in Project does not support the control for report theme colors.


## Example

In the following example, the chart is the first shape in the "Simple scalar chart" report. The example sets the chart color scheme to monochromatic greens.


```vb
Sub SetChartColor()
    Dim chartShape As Shape
    
    Set chartShape = ActiveProject.Reports("Simple scalar chart").Shapes(1)
    
    ' ChartColor values 10 - 26 correspond to the Change Colors menu
    ' on the DESIGN tab of the CHART TOOLS ribbon.
    chartShape.Chart.ChartColor = 26
End Sub
```


## Property value

 **VARIANT**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
