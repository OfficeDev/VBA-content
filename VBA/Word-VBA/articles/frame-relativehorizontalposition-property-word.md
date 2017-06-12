---
title: Frame.RelativeHorizontalPosition Property (Word)
keywords: vbawd10.chm153747463
f1_keywords:
- vbawd10.chm153747463
ms.prod: word
api_name:
- Word.Frame.RelativeHorizontalPosition
ms.assetid: ff95768c-26c1-4be2-0a64-8626f2241e2a
ms.date: 06/08/2017
---


# Frame.RelativeHorizontalPosition Property (Word)

Specifies the relative horizontal position of a frame. Read/write  **[WdRelativeHorizontalPosition](wdrelativehorizontalposition-enumeration-word.md)** .


## Syntax

 _expression_ . **RelativeHorizontalPosition**

 _expression_ An expression that represents a **[Frame](frame-object-word.md)** object.


## Example

This example adds a frame around the selection and aligns the frame horizontally with the right margin.


```vb
Set myFrame = ActiveDocument.Frames.Add(Range:=Selection.Range) 
With myFrame 
 .RelativeHorizontalPosition = _ 
 wdRelativeHorizontalPositionMargin 
 .HorizontalPosition = wdFrameRight 
End With
```




## See also


#### Concepts


[Frame Object](frame-object-word.md)

