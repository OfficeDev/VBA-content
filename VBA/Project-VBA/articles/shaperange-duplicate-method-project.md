---
title: ShapeRange.Duplicate Method (Project)
ms.prod: project-server
ms.assetid: c7af186e-616a-f20c-e2f3-8b0319e6af44
ms.date: 06/08/2017
---


# ShapeRange.Duplicate Method (Project)
Duplicates a shape range and returns a reference to the copy.

## Syntax

 _expression_. **Duplicate**

 _expression_ A variable that represents a **ShapeRange** object.


### Return value

 **ShapeRange**


### Remarks

The horizontal offset and vertical offset of the duplicated shape range are both 12 points from the original shape range.


### Example

The following example uses the report created by the code example in the  **[ShapeRange.Apply](shaperange-apply-method-project.md)** method. The example duplicates a shape range that contains two shapes, and then vertically flips and selects the new shape range.


```vb
Sub DuplicateShapeRange()
    Dim theReport As Report
    Dim shp1 As shape
    Dim shp2 As shape
    Dim shp3 As shape
    Dim reportName As String
    Dim sRange1 As ShapeRange
    Dim sRange2 As ShapeRange
    
    reportName = "Apply Report"
    
    Set theReport = ActiveProject.Reports(reportName)
    Set shp1 = theReport.Shapes(1)
    Set shp2 = theReport.Shapes(2)
    Set shp3 = theReport.Shapes(3)
    
    Set sRange1 = theReport.Shapes.Range(Array(2, 3))
    
    Set sRange2 = sRange1.Duplicate()
    
    sRange2.Flip msoFlipVertical
    sRange2.Select
End Sub
```


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
[Shape.Duplicate Method](shape-duplicate-method-project.md)
