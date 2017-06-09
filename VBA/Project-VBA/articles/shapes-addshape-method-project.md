---
title: Shapes.AddShape Method (Project)
ms.prod: project-server
ms.assetid: 58af0a51-a455-5c9a-1cae-e56dc67a08a5
ms.date: 06/08/2017
---


# Shapes.AddShape Method (Project)
Adds a shape of the specified AutoShape type to a report, and returns a  **Shape** object that represents the new shape.

## Syntax

 _expression_. **AddShape** _(Type,_ _Left,_ _Top,_ _Width,_ _Height)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**MsoAutoShapeType**|Specifies the type of AutoShape to create.|
| _Left_|Required|**Single**|The position, in points, of the left edge of the AutoShape.|
| _Top_|Required|**Single**|The position, in points, of the top edge of the AutoShape.|
| _Width_|Required|**Single**|The width, in points, of the AutoShape.|
| _Height_|Required|**Single**|The height, in points, of the AutoShape.|
| _Type_|Required|MSOAUTOSHAPETYPE||
| _Left_|Required|FLOAT||
| _Top_|Required|FLOAT||
| _Width_|Required|FLOAT||
| _Height_|Required|FLOAT||
|Name|Required/Optional|Data type|Description|

### Return value

 **Shape**


## Remarks

To change the type of an AutoShape, set the  **AutoShapeType** property.


## Example

The following example creates a report that contains two cloud shapes, and then changes the second cloud shape to a yellow speech balloon.


```vb
Sub TestShapes()
    Dim shapeReport As Report
    Dim reportName As String
    
    ' Add a report.
    reportName = "Shape report"
    Set shapeReport = ActiveProject.Reports.Add(reportName)

    ' Add two clouds.
    Dim cloudShape1 As shape
    Dim cloudShape2 As shape
    Set cloudShape1 = shapeReport.Shapes.AddShape(msoShapeCloud, 20, 20, 100, 60)
    Set cloudShape2 = shapeReport.Shapes.AddShape(msoShapeCloud, 100, 200, 60, 100)
    
    ' Change the blue cloud to a yellow speech balloon.
    cloudShape2.AutoShapeType = msoShapeBalloon
    cloudShape2.Fill.ForeColor.RGB = &;H80FFFF
End Sub
```


## See also


#### Other resources


[Shapes Object](shapes-object-project.md)
[Shape Object](shape-object-project.md)
[AutoShapeType Property](shape-autoshapetype-property-project.md)
[MsoAutoShapeType Enumeration (Office)](http://msdn.microsoft.com/en-us/library/office/ff862770%28v=office.15%29)
