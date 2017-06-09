---
title: Shapes.AddLine Method (Project)
ms.prod: project-server
ms.assetid: 697a5972-4b24-8e77-b42f-b064019906fa
ms.date: 06/08/2017
---


# Shapes.AddLine Method (Project)
Adds a line to a report, and returns a  **Shape** object that represents the line.

## Syntax

 _expression_. **AddLine** _(BeginX,_ _BeginY,_ _EndX,_ _EndY)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _BeginX_|Required|**Single**|The horizontal position (in points) of the starting point, relative to the left edge of the report.|
| _BeginY_|Required|**Single**|The vertical position (in points) of the starting point, relative to the top edge of the report.|
| _EndX_|Required|**Single**|The horizontal position (in points) of the end point, relative to the left edge of the report.|
| _EndY_|Required|**Single**|The vertical position (in points) of the end point, relative to the top edge of the report.|
| _BeginX_|Required|FLOAT||
| _BeginY_|Required|FLOAT||
| _EndX_|Required|FLOAT||
| _EndY_|Required|FLOAT||
|Name|Required/Optional|Data type|Description|

### Return value

 **Shape**


## Remarks

To format the line, use the  **Shape.Line** property.


## Example

The following example creates a violet dashed line with an arrow at the end.


```vb
Sub AddBigArrow()
    Dim shapeReport As Report
    Dim reportName As String
    Dim lineShape As shape
    
    ' Add a report.
    reportName = "Line report"
    Set shapeReport = ActiveProject.Reports.Add(reportName)

    Set lineShape = shapeReport.Shapes.AddLine(20, 50, 320, 100)
    
    With lineShape.Line
        .DashStyle = msoLineDashDot
        .Weight = 3
        .EndArrowheadStyle = msoArrowheadTriangle
        .EndArrowheadWidth = msoArrowheadWidthMedium
        .ForeColor.RGB = &;HFF0090
    End With
End Sub
```


## See also


#### Other resources


[Shapes Object](shapes-object-project.md)
[Shape Object](shape-object-project.md)
[Line Property](shape-line-property-project.md)
