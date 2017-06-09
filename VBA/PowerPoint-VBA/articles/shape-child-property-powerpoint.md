---
title: Shape.Child Property (PowerPoint)
keywords: vbapp10.chm547066
f1_keywords:
- vbapp10.chm547066
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Child
ms.assetid: 53371144-eabb-3f1f-f9cf-9a4e7b701d5f
ms.date: 06/08/2017
---


# Shape.Child Property (PowerPoint)

 **MsoTrue** if the shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only.


## Syntax

 _expression_. **Child**

 _expression_ A variable that represents a **Shape** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Child** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The shape is not a child shape or, if a shape range, all child shapes do not belong to the same parent.|
|**msoTrue**| The shape is a child shape or, if a shape range, all child shapes belong to the same parent.|

## Example

This example selects the first shape in the canvas, and if the selected shape is a child shape, fills the shape with the specified color. This example assumes that the first shape in the active presentation is a drawing canvas that contains multiple shapes.


```vb
Sub FillChildShape()



    'Select the first shape in the drawing canvas

    ActivePresentation.Slides(1).Shapes(1).CanvasItems(1).Select



    'Fill selected shape if it is a child shape

    With ActiveWindow.Selection



        If .ShapeRange.Child = msoTrue Then

            .ShapeRange.Fill.ForeColor.RGB = RGB(Red:=100, Green:=0, Blue:=200)

        Else

            MsgBox "This shape is not a child shape."

        End If



    End With



End Sub
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

