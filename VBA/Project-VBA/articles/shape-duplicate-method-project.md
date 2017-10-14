---
title: Shape.Duplicate Method (Project)
ms.prod: project-server
ms.assetid: 19917b35-589e-1cd3-e9eb-5efa13e02793
ms.date: 06/08/2017
---


# Shape.Duplicate Method (Project)
Duplicates a shape and returns a reference to the copy.

## Syntax

 _expression_. **Duplicate**

 _expression_ A variable that represents a **Shape** object.


### Return value

 **Shape**


## Example

The following example uses the report created by the code example in the  **[Shape.Apply](shape-apply-method-project.md)** method. The example duplicates a shape, and then rotates, horizontally flips, and selects the new shape. The horizontal offset and vertical offset of the new shape are both 12 points.


```vb
Sub DuplicateShape()
    Dim theReport As Report
    Dim shp1 As shape
    Dim duplicatedShape As shape
    Dim reportName As String
    
    reportName = "Apply Report"
    
    Set theReport = ActiveProject.Reports(reportName)
    Set shp1 = theReport.Shapes(1)
    
    Set duplicatedShape = shp1.Duplicate
    
    pos1 = shp1.left
    pos2 = duplicatedShape.left
    Debug.Print "Horizontal offset: " &; CStr(pos2 - pos1)
    
    pos1 = shp1.top
    pos2 = duplicatedShape.top
    Debug.Print "Vertical offset: " &; CStr(pos2 - pos1)
   
    duplicatedShape.Rotation = 30
    duplicatedShape.Flip msoFlipHorizontal
    
    duplicatedShape.Select
End Sub
```


## See also


#### Other resources


[Shape Object](shape-object-project.md)
[ShapeRange.Duplicate Method](shaperange-duplicate-method-project.md)
