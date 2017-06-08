---
title: ShadowFormat.OffsetX Property (Publisher)
keywords: vbapb10.chm3670274
f1_keywords:
- vbapb10.chm3670274
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OffsetX
ms.assetid: 2b34ace8-5c3b-002b-df96-43c8aef2fbd2
ms.date: 06/08/2017
---


# ShadowFormat.OffsetX Property (Publisher)

Returns or sets a  **Variant** value indicating the vertical offset of the shadow from the specified shape. A positive value offsets the shadow below the shape; a negative value offsets it above the shape. Read/write.


## Syntax

 _expression_. **OffsetX**

 _expression_A variable that represents an  **ShadowFormat** object.


### Return Value

Variant


## Remarks

Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

If you want to nudge a shadow horizontally or vertically from its current position without having to specify an absolute position, use the  **[IncrementOffsetX](shadowformat-incrementoffsetx-method-publisher.md)** method or the  **[IncrementOffsetY](shadowformat-incrementoffsety-method-publisher.md)** method.


## Example

This example sets the horizontal and vertical offsets of the shadow for shape three on page one of the active publication. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it.


```vb
With ActiveDocument.Pages(1).Shapes(3).Shadow 
 .Visible = True 
 .OffsetX = 5 
 .OffsetY = -3 
End With
```


