---
title: ShadowFormat.OffsetX Property (Word)
keywords: vbawd10.chm164364390
f1_keywords:
- vbawd10.chm164364390
ms.prod: word
api_name:
- Word.ShadowFormat.OffsetX
ms.assetid: 5556921b-b96b-7e28-8cd4-7be3475f6a6f
ms.date: 06/08/2017
---


# ShadowFormat.OffsetX Property (Word)

Returns or sets the horizontal offset (in points) of the shadow from the specified shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left. Read/write  **Single** .


## Syntax

 _expression_ . **OffsetX**

 _expression_ An expression that returns a **[ShadowFormat](shadowformat-object-word.md)** object.


## Remarks

If you want to nudge a shadow horizontally or vertically from its current position without having to specify an absolute position, use the  **[IncrementOffsetX](shadowformat-incrementoffsetx-method-word.md)** or **[IncrementOffsetY](shadowformat-incrementoffsety-method-word.md)** method.


## Example

This example sets the horizontal and vertical offsets for the shadow of shape three on myDocument. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(3).Shadow 
 .Visible = True 
 .OffsetX = 5 
 .OffsetY = -3 
End With
```


## See also


#### Concepts


[ShadowFormat Object](shadowformat-object-word.md)

