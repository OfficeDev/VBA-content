---
title: ShapeRange.ScaleWidth Method (Project)
ms.prod: project-server
ms.assetid: 6087bb9c-c111-7f2e-95d9-334af18fe37d
ms.date: 06/08/2017
---


# ShapeRange.ScaleWidth Method (Project)
Scales the width of the range of shapes by a specified factor.

## Syntax

 _expression_. **ScaleWidth** _(Factor,_ _RelativeToOriginalSize,_ _fScale)_

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Factor_|Required|**Single**|The ratio between the width of the shape after you resize it and the current width. For example, to make a rectangle 50 percent wider, specify 1.5 for the  _Factor_ parameter.|
| _RelativeToOriginalSize_|Required|**[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**|**msoFalse** scales each shape relative to its current size. For Project, the value must be **msoFalse**.|
| _fScale_|Optional|**[MsoScaleFrom](http://msdn.microsoft.com/en-us/library/office/ff863348%28v=office.15%29)**|Specifies which part of the shape retains its position when the shape is scaled.|
| _Factor_|Required|FLOAT||
| _RelativeToOriginalSize_|Required|MSOTRISTATE||
| _fScale_|Optional|MSOSCALEFROM||

### Return value

 **Nothing**


## Remarks

A  _RelativeToOriginalSize_ parameter value of **msoTrue** scales a shape relative to its original size, which applies only to a picture or OLE object.


## Example

The following example creates two cylindrical shapes, assigns them to a shape range, and then scales the shapes in height and width. If you set a breakpoint on the first  **ScaleHeight** statement, you can step through the code and see the changes from scaling and from using the _fScale_ parameter.


```vb
Sub ScaleShapes()
    Dim theReport As Report
    Dim shp1 As shape
    Dim shp2 As shape
    Dim reportName As String
    Dim sRange As ShapeRange
    
    reportName = "Scale Report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set shp1 = theReport.Shapes.AddShape(msoShapeCan, 20, 50, 20, 30)
    Set shp2 = theReport.Shapes.AddShape(msoShapeCan, 140, 50, 30, 50)
        
    Set sRange = theReport.Shapes.Range(Array(1, 2))
    sRange.ScaleHeight 2, msoFalse
    sRange.ScaleWidth 2, msoFalse

    sRange.ScaleHeight 2, msoFalse, msoScaleFromMiddle
    sRange.ScaleWidth 2, msoFalse, msoScaleFromTopLeft
End Sub
```


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)
[MsoScaleFrom](http://msdn.microsoft.com/en-us/library/office/ff863348%28v=office.15%29)
