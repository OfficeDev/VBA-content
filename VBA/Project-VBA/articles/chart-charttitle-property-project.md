---
title: Chart.ChartTitle Property (Project)
ms.prod: project-server
ms.assetid: eb2e9c18-1dcc-3d66-e73c-b5d0dfa88472
ms.date: 06/08/2017
---


# Chart.ChartTitle Property (Project)
Gets an  **Office.IMsoChartTitle** object that represents the title of the specified chart. Read-only **IMsoChartTitle**.

## Syntax

 _expression_. **ChartTitle**

 _expression_ A variable that represents a **Chart** object.


## Remarks

To manually edit the text of a chart title, click in the title area. To change the title format, select the chart, and then, on the ribbon under  **CHART TOOLS**, choose the  **FORMAT** tab.


## Example

The following example changes the chart title and sets the title above the chart.


```vb
Sub ChangeChartTitle()
    Dim chartShape As Shape
    
    Set chartShape = ActiveProject.Reports("Simple scalar chart").Shapes(1)
    
    With chartShape.Chart
        If Not .HasTitle Then
            .HasTitle = True
        End If
        
        .ChartTitle.Text = "New chart title"
        .SetElement (msoElementChartTitleAboveChart)
    End With
End Sub
```


## Property value

 **IMSOCHARTTITLE**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
