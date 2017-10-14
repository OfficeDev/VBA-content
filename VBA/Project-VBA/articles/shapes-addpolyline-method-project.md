---
title: Shapes.AddPolyline Method (Project)
ms.prod: project-server
ms.assetid: c61cbaf3-b687-b137-e4a2-8f9061dfc0f0
ms.date: 06/08/2017
---


# Shapes.AddPolyline Method (Project)
Creates an open polyline or a closed polygon drawing, and returns a  **Shape** object that represents the new polyline or polygon.

## Syntax

 _expression_. **AddPolyline** _(SafeArrayOfPoints)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SafeArrayOfPoints_|Required|**Variant**|An array of coordinate pairs that specifies the vertices of the polyline.|
| _SafeArrayOfPoints_|Required|VARIANT||
|Name|Required/Optional|Data type|Description|

### Return value

 **Shape**


## Remarks

To form a closed polygon, assign the same coordinates to the first and last vertices in the polyline drawing. For a closed polygon, the default shape fill color is a medium blue: &;HD59B5B, or  `RGB(Red:=91, Green:=155, Blue:=213)`.


## Example

Because the first and last points are the same, the following example creates a closed octagon. The violet line is two points wide; the octagon is filled with a gold color.


```vb
Sub AddOctagon()
    Dim shapeReport As Report
    Dim reportName As String
    Dim polylineShape As shape
    
    ' Add a report.
    reportName = "Polyline report"
    Set shapeReport = ActiveProject.Reports.Add(reportName)
    
    Dim octArray(1 To 9, 1 To 2) As Single
    octArray(1, 1) = 9
    octArray(1, 2) = 8
    octArray(2, 1) = 19
    octArray(2, 2) = 8
    octArray(3, 1) = 26
    octArray(3, 2) = 15
    octArray(4, 1) = 26
    octArray(4, 2) = 25
    octArray(5, 1) = 19
    octArray(5, 2) = 32
    octArray(6, 1) = 9
    octArray(6, 2) = 32
    octArray(7, 1) = 2
    octArray(7, 2) = 25
    octArray(8, 1) = 2
    octArray(8, 2) = 15
    octArray(9, 1) = 9
    octArray(9, 2) = 8
    
    Set polylineShape = shapeReport.Shapes.AddPolyline(octArray)
    
    With polylineShape.Line
        .Weight = 2
        .ForeColor.RGB = &;HFF0090    ' Violet color.
    End With
    
    polylineShape.Fill.ForeColor.RGB = &;H10D0D0    ' Gold color.
End Sub
```


## See also


#### Other resources


[Shapes Object](shapes-object-project.md)
[Shape Object](shape-object-project.md)
[Line Property](shape-line-property-project.md)
[Fill Property](shape-fill-property-project.md)
