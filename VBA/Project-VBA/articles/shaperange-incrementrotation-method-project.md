---
title: ShapeRange.IncrementRotation Method (Project)
ms.prod: project-server
ms.assetid: 404bd4de-7c5f-3107-baa1-63c4c3362537
ms.date: 06/08/2017
---


# ShapeRange.IncrementRotation Method (Project)
Rotates each shape in the shape range around the z-axis by the specified number of degrees.

## Syntax

 _expression_. **IncrementRotation** _(Increment)_

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|The number of degrees each shape is to be rotated. A positive value rotates the shapes clockwise; a negative value rotates the shapes counterclockwise.|
| _Increment_|Required|FLOAT||
|Name|Required/Optional|Data type|Description|

### Return value

 **Nothing**


## Remarks

The  _Increment_ parameter can be a value from -3600 to 3600.


## Example

The following example shows the difference between rotating a shape and rotating a shape range. The example creates a shape range that contains two cylinders, rotates the shape range 30 degrees clockwise, and then rotates the second shape in the range 30 degrees counterclockwise. If you set a breakpoint on the last  **IncrementRotation** statement, and then step through the code, you can see how the rotation works.


```vb
Sub RotateShapes()
    Dim theReport As Report
    Dim shp1 As shape
    Dim shp2 As shape
    Dim shpGroup As shape
    Dim reportName As String
    Dim sRange1 As ShapeRange
    
    reportName = "Rotate Report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set shp1 = theReport.Shapes.AddShape(msoShapeCan, 20, 30, 100, 100)
    Set shp2 = theReport.Shapes.AddShape(msoShapeCan, 140, 30, 100, 100)
        
    Set sRange1 = theReport.Shapes.Range(Array(1, 2))
    sRange1.IncrementRotation 30

    sRange1(2).IncrementRotation -30
End Sub
```


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
