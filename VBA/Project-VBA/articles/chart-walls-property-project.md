---
title: Chart.Walls Property (Project)
ms.prod: project-server
ms.assetid: 8404e5cb-8da2-49b4-c49a-488d67457681
ms.date: 06/08/2017
---


# Chart.Walls Property (Project)
Gets an  **Office.IMsoWalls** object that represents the walls of a 3-D chart. Read-only **IMsoWalls**.

## Syntax

 _expression_. **Walls**

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _fBackWall_|Optional|**Boolean**|Default value =  **True**. The  _fBackWall_ parameter has no effect in Project.|

## Example

The following example sets the wall borders of the 3-D chart to a red line that is three points wide.


```vb
Sub FormatWalls()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3-D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    With chartShape.Chart.Walls.Border
        .Weight = 3
        .Color = &;HFF
    End With
End Sub
```


## Property value

 **IMSOWALLS**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
