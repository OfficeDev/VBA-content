---
title: ShapeRange.RelativeHorizontalPosition Property (Word)
keywords: vbawd10.chm162857260
f1_keywords:
- vbawd10.chm162857260
ms.prod: word
api_name:
- Word.ShapeRange.RelativeHorizontalPosition
ms.assetid: f1150705-3004-3987-3826-70f402105a99
ms.date: 06/08/2017
---


# ShapeRange.RelativeHorizontalPosition Property (Word)

Specifies the relative horizontal position of a range of shapes. Read/write  **[WdRelativeHorizontalPosition](wdrelativehorizontalposition-enumeration-word.md)** .


## Syntax

 _expression_ . **RelativeHorizontalPosition**

 _expression_ An expression that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Example

This example repositions the selected shape object.


```vb
With Selection.ShapeRange 
 .Left = InchesToPoints(0.6) 
 .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage 
 .Top = InchesToPoints(1) 
 .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph 
End With
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

