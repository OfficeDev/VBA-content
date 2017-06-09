---
title: Shape.ZOrder Method (PowerPoint)
keywords: vbapp10.chm547014
f1_keywords:
- vbapp10.chm547014
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.ZOrder
ms.assetid: 3317b5c3-611f-7cf8-3ce3-6d09255aa75f
ms.date: 06/08/2017
---


# Shape.ZOrder Method (PowerPoint)

Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).


## Syntax

 _expression_. **ZOrder**( **_ZOrderCmd_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|Required|**MsoZOrderCmd**|Specifies where to move the specified shape relative to the other shapes.|

## Remarks

The  _ZOrderCmd_ parameter value can be one of these **MsoZOrderCmd** constants.


||
|:-----|
|**msoBringForward**|
|**msoBringInFrontOfText**|
|**msoBringToFront**|
|**msoSendBackward**|
|**msoSendBehindText**|
|**msoSendToBack**|
The  **msoBringInFrontOfText** and **msoSendBehindText** constants should be used only in Microsoft Office Word.

Use the  **ZOrderPosition** property to determine a shape's current position in the z-order.


## Example

This example adds an oval to  `myDocument` and then places the oval second from the back in the z-order if there is at least one other shape on the slide.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300)

    While .ZOrderPosition > 2

        .ZOrder msoSendBackward

    Wend

End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

