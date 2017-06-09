---
title: ShapeRange.TopRelative Property (Word)
keywords: vbawd10.chm162857161
f1_keywords:
- vbawd10.chm162857161
ms.prod: word
api_name:
- Word.ShapeRange.TopRelative
ms.assetid: 6162d05b-0610-7a6b-0224-7bd6f658276b
ms.date: 06/08/2017
---


# ShapeRange.TopRelative Property (Word)

Returns or sets a  **Single** that represents the relative top position of a range of shapes. Read/write.


## Syntax

 _expression_ . **TopRelative**

 _expression_ An expression that returns a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

Use this property with the  **[RelativeVerticalPosition](shaperange-relativeverticalposition-property-word.md)** property. When set to **wdShapePositionRelativeNone** (-999999) (see the **[WdShapePositionRelative](wdshapepositionrelative-enumeration-word.md)** enumeration), this property should be ignored because the shape does not use percent positioning. The vertical position is solely determined by the **[Top](shaperange-top-property-word.md)** property.


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

