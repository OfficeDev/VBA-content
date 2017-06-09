---
title: LineFormat.InsetPen Property (PowerPoint)
keywords: vbapp10.chm553016
f1_keywords:
- vbapp10.chm553016
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.InsetPen
ms.assetid: 07a69459-0a24-c9b8-5aba-103b39d8b1af
ms.date: 06/08/2017
---


# LineFormat.InsetPen Property (PowerPoint)

Detemines whether to draw lines on the inside of a specified shape. Read/write.


## Syntax

 _expression_. **InsetPen**

 _expression_ A variable that represents an **LineFormat** object.


### Return Value

MsoTriState


## Remarks

An error occurs if this property attempts to set an inset pen drawing on any Microsoft Office AutoShape that does not support inset pen drawing.

The value of the  **InsetPen** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. An inset pen is not enabled.|
|**msoTrue**| An inset pen is enabled.|

## Example

The following line of code enables an inset pen for a shape. This example assumes that the first slide of the active presentation contains a shape and that the shape supports inset pen drawing.


```vb
Sub DrawLinesInsideShape

    ActivePresentation.Slides(1).Shapes(1).Line.InsetPen = msoTrue

End Sub
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-powerpoint.md)

