---
title: Shape.ScaleHeight Method (Project)
ms.prod: project-server
ms.assetid: 9dd7a3ee-9149-4aa2-c4ad-5fc0ca45ee21
ms.date: 06/08/2017
---


# Shape.ScaleHeight Method (Project)
Scales the height of the shape by a specified factor.

## Syntax

 _expression_. **ScaleHeight** _(Factor,_ _RelativeToOriginalSize,_ _fScale)_

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Factor_|Required|**Single**|The ratio between the height of the shape after you resize it and the current height. For example, to make a rectangle 50 percent higher, specify 1.5 for the  _Factor_ parameter.|
| _RelativeToOriginalSize_|Required|MSOTRISTATE|**msoFalse** scales the shape relative to its current size. For Project, the value must be **msoFalse**.|
| _fScale_|Optional|MSOSCALEFROM|Specifies which part of the shape retains its position when the shape is scaled.|
| _Factor_|Required|FLOAT||
| _RelativeToOriginalSize_|Required|MSOTRISTATE||
| _fScale_|Optional|MSOSCALEFROM||
|Name|Required/Optional|Data type|Description|

### Return value

 **Nothing**


## Remarks

A  _RelativeToOriginalSize_ parameter value of **msoTrue** scales a shape relative to its original size, which applies only to a picture or OLE object.


## Example

The following example creates two cylindrical shapes, and then scales the first shape in height and width. If you set a breakpoint on the first  **ScaleHeight** statement, you can step through the code and see the changes from scaling and from using the _fScale_ parameter.


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
    
    shp1.ScaleHeight 2, msoFalse
    shp1.ScaleWidth 2, msoFalse

    shp1.ScaleHeight 2, msoFalse, msoScaleFromMiddle
    shp1.ScaleWidth 2, msoFalse, msoScaleFromTopLeft
End Sub
```


## See also


#### Other resources


[Shape Object](shape-object-project.md)
[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)
[MsoScaleFrom](http://msdn.microsoft.com/en-us/library/office/ff863348%28v=office.15%29)
