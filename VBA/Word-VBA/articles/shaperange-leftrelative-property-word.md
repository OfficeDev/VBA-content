---
title: ShapeRange.LeftRelative Property (Word)
keywords: vbawd10.chm162857160
f1_keywords:
- vbawd10.chm162857160
ms.prod: word
api_name:
- Word.ShapeRange.LeftRelative
ms.assetid: c253909c-2204-6f38-d7d3-8a0587842718
ms.date: 06/08/2017
---


# ShapeRange.LeftRelative Property (Word)

Returns or sets a  **Single** that represents the relative left position of a range of shapes. Read/write.


## Syntax

 _expression_ . **LeftRelative**

 _expression_ An expression that returns a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

Use this property with the  **[RelativeHorizontalPosition](shaperange-relativehorizontalposition-property-word.md)** property. When set to **wdShapePositionRelativeNone** (-999999) (see the **[WdShapePositionRelative](wdshapepositionrelative-enumeration-word.md)** enumeration), this property should be ignored because the shape does not use percent positioning. The horizontal position is solely determined by the **[Left](shaperange-left-property-word.md)** property.


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

