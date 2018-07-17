---
title: ShapeRange.Apply Method (Project)
ms.prod: project-server
ms.assetid: 5b100f4a-99a0-77f2-772a-203b2f836293
ms.date: 06/08/2017
---


# ShapeRange.Apply Method (Project)
Applies formatting to a shape range, where the formatting information has been copied by using the  **[PickUp](shape-pickup-method-project.md)** method.

## Syntax

 _expression_. **Apply**

 _expression_ A variable that represents a **ShapeRange** object.


### Return value

 **Nothing**


## Example

The following example creates three cylindrical shapes, colors the first shape red, adds the second and third shapes to a shape range, copies the formatting of the first shape, and then applies the formatting to the shape range.


```vb
Sub ApplyShapeFormat()
    Dim theReport As Report
    Dim shp1 As shape
    Dim shp2 As shape
    Dim shp3 As shape
    Dim reportName As String
    Dim sRange As ShapeRange
    
    reportName = "Apply Report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set shp1 = theReport.Shapes.AddShape(msoShapeCan, 10, 30, 100, 100)
    shp1.Name = "Shape 1"
    shp1.Fill.ForeColor.RGB = &;H1010FF  ' Red color.
    
    ' Blue default color.
    Set shp2 = theReport.Shapes.AddShape(msoShapeCan, 30, 140, 100, 100)
    
    ' Blue default color.
    Set shp3 = theReport.Shapes.AddShape(msoShapeCan, 140, 140, 100, 100)
    
    Set sRange = theReport.Shapes.Range(Array(2, 3))
    
    theReport.Shapes("Shape 1").PickUp
    sRange.Apply
End Sub
```


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
[PickUp Method](shape-pickup-method-project.md)
[Shape.Apply Method](shape-object-project.md)
