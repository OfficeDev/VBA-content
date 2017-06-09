---
title: ShapeRange.ZOrder Method (PowerPoint)
keywords: vbapp10.chm548014
f1_keywords:
- vbapp10.chm548014
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.ZOrder
ms.assetid: 906620bd-9293-694a-002d-97e760de988a
ms.date: 06/08/2017
---


# ShapeRange.ZOrder Method (PowerPoint)

Moves the specified shape range in front of or behind other shapes in the collection (that is, changes the shape range's position in the z-order).


## Syntax

 _expression_. **ZOrder**( **_ZOrderCmd_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|Required|**MsoZOrderCmd**|Specifies where to move the specified shape range relative to the other shapes.|

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


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

