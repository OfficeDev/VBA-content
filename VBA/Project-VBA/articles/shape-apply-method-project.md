---
title: Shape.Apply Method (Project)
ms.prod: project-server
ms.assetid: 8d7a29f0-6a69-f643-6726-0c85247fb957
ms.date: 06/08/2017
---


# Shape.Apply Method (Project)
Applies formatting to a shape, where the formatting information has been copied by using the  **[PickUp](shape-pickup-method-project.md)** method.

## Syntax

 _expression_. **Apply**

 _expression_ A variable that represents a **Shape** object.


### Return value

 **Nothing**


## Example

The following example creates two cylindrical shapes, colors the first shape red, copies the formatting of the first shape, and then applies it to the second shape.


```vb
Sub ApplyShapeFormat()
    Dim theReport As Report
    Dim shp1 As shape
    Dim shp2 As shape
    Dim reportName As String
    Dim sRange As ShapeRange
    
    reportName = "Apply Report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set shp1 = theReport.Shapes.AddShape(msoShapeCan, 10, 30, 100, 100)
    shp1.Name = "Shape 1"
    shp1.Fill.ForeColor.RGB = &;H1010FF  ' Red color.
    
    Set shp2 = theReport.Shapes.AddShape(msoShapeCan, 30, 140, 100, 100)
    shp2.Name = "Shape 2"               ' Blue default color.
    
    With theReport
        .Shapes("Shape 1").PickUp
        .Shapes("Shape 2").Apply
    End With
End Sub
```


## See also


#### Other resources


[Shape Object](shape-object-project.md)
[PickUp Method](shape-pickup-method-project.md)
[ShapeRange.Apply Method](shaperange-apply-method-project.md)
