---
title: Window.RangeFromPoint Method (Word)
keywords: vbawd10.chm157417582
f1_keywords:
- vbawd10.chm157417582
ms.prod: word
api_name:
- Word.Window.RangeFromPoint
ms.assetid: 27c6ed94-0b47-3e0d-701f-09e72b115910
ms.date: 06/08/2017
---


# Window.RangeFromPoint Method (Word)

Returns the  **Range** or **Shape** object that is located at the point specified by the screen position coordinate pair.


## Syntax

 _expression_ . **RangeFromPoint**( **_x_** , **_y_** )

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x_|Required| **Long**|The horizontal distance (in pixels) from the left edge of the screen to the point.|
| _y_|Required| **Long**|The vertical distance (in pixels) from the top of the screen to the point.|

### Return Value

Object


## Remarks

If no range or shape is located at the coordinate pair specified, the method returns  **Nothing** .


## Example

This example creates a new document and adds a five-point star. It then obtains the screen location of the shape and calculates where the center of the shape is. Using these coordinates, the example uses the  **RangeFromPoint** method to return a reference to the shape and change its fill color.


```vb
Dim pLeft As Long 
Dim pTop As Long 
Dim pWidth As Long 
Dim pHeight As Long 
Dim newShape As Object 
Dim newDoc As New Document 
 
With newDoc 
 .Shapes.AddShape msoShape5pointStar, _ 
 288, 100, 100, 72 
 .ActiveWindow.GetPoint pLeft, pTop, _ 
 pWidth, pHeight, .Shapes(1) 
 Set newShape = .ActiveWindow.RangeFromPoint(pLeft _ 
 + pWidth * 0.5, pTop + pHeight * 0.5) 
 newShape.Fill.ForeColor.RGB = RGB(80, 160, 130) 
End With
```


## See also


#### Concepts


[Window Object](window-object-word.md)

