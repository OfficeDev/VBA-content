---
title: Chart.SideWall Property (Project)
keywords: vbapj.chm131633
f1_keywords:
- vbapj.chm131633
ms.prod: project-server
ms.assetid: d8b74dc2-7a22-1064-972d-876396414e2c
ms.date: 06/08/2017
---


# Chart.SideWall Property (Project)
Gets an  **Office.IMsoWalls** object that allows the user to individually format the side wall of a 3-D chart. Read-only **IMsoWalls**.

## Syntax

 _expression_. **SideWall**

 _expression_ A variable that represents a **Chart** object.


## Example

The following example colors the side wall of the 3-D chart blue. In Project, red is the last byte of a hexadecimal value.


```vb
Sub FormatSideWall()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3-D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.SideWall.Fill.ForeColor.RGB = &;HFF0000
End Sub
```


## Property value

 **IMSOWALLS**


## See also


#### Other resources


[Chart Object](chart-object-project.md)
