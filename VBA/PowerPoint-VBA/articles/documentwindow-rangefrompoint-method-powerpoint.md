---
title: DocumentWindow.RangeFromPoint Method (PowerPoint)
keywords: vbapp10.chm511026
f1_keywords:
- vbapp10.chm511026
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.RangeFromPoint
ms.assetid: 74bc61e5-6c6d-0510-b549-e325dd67c7a7
ms.date: 06/08/2017
---


# DocumentWindow.RangeFromPoint Method (PowerPoint)

Returns the  **Shape** object that is located at the point specified by the screen position coordinate pair. If no shape is located at the coordinate pair specified, then the method returns **Nothing**.


## Syntax

 _expression_. **RangeFromPoint**( **_x_**, **_y_** )

 _expression_ A variable that represents a **DocumentWindow** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x_|Required|**Long**|The horizontal distance (in pixels) from the left edge of the screen to the point.|
| _y_|Required|**Long**|The vertical distance (in pixels) from the top of the screen to the point.|

## Example

This example adds a new five-point star to slide one using the coordinates (288, 100). It then converts those coordinates from points to pixels, uses the  **RangeFromPoint** method to return a reference to the new object, and changes the fill color of the star.


```vb
Dim myPointX As Integer, myPointY As Integer
Dim myShape As Object

ActivePresentation.Slides(1).Shapes _
    .AddShape(msoShape5pointStar, 288, 100, 100, 72).Select

myPointX = ActiveWindow.PointsToScreenPixelsX(288)
myPointY = ActiveWindow.PointsToScreenPixelsY(100)
Set myShape = ActiveWindow.RangeFromPoint(myPointX, myPointY)
myShape.Fill.ForeColor.RGB = RGB(80, 160, 130)
```


## See also


#### Concepts


[DocumentWindow Object](documentwindow-object-powerpoint.md)


