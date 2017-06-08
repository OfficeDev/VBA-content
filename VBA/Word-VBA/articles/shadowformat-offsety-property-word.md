---
title: ShadowFormat.OffsetY Property (Word)
keywords: vbawd10.chm164364391
f1_keywords:
- vbawd10.chm164364391
ms.prod: word
api_name:
- Word.ShadowFormat.OffsetY
ms.assetid: 3ec87258-bcab-42fe-cc58-2d59b5dd4d65
ms.date: 06/08/2017
---


# ShadowFormat.OffsetY Property (Word)

Returns or sets the vertical offset (in points) of the shadow from the specified shape. Read/write  **Single** .


## Syntax

 _expression_ . **OffsetY**

 _expression_ An expression that returns a **[ShadowFormat](shadowformat-object-word.md)** object.


## Remarks

A positive value offsets the shadow below the shape; a negative value offsets it above the shape. If you want to nudge a shadow horizontally or vertically from its current position without having to specify an absolute position, use the  **[IncrementOffsetX](shadowformat-incrementoffsetx-method-word.md)** or **[IncrementOffsetY](shadowformat-incrementoffsety-method-word.md)** method.


## Example

This example sets the horizontal and vertical offsets for the shadow of shape three in myDocument. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it.


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

