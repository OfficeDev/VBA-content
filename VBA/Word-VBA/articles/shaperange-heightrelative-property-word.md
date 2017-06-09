---
title: ShapeRange.HeightRelative Property (Word)
keywords: vbawd10.chm162857163
f1_keywords:
- vbawd10.chm162857163
ms.prod: word
api_name:
- Word.ShapeRange.HeightRelative
ms.assetid: f0414af1-f09a-475d-5e96-bfe2ceab8901
ms.date: 06/08/2017
---


# ShapeRange.HeightRelative Property (Word)

Returns or sets a  **Single** that represents the percentage of the target shape to which the range of shapes is sized. Read/write.


## Syntax

 _expression_ . **HeightRelative**

 _expression_ An expression that returns a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

Use this property with the  **[RelativeVerticalSize](shaperange-relativeverticalsize-property-word.md)** property. When set to **wdShapeSizeRelativeNone** (-999999) (see the **[WdShapeSizeRelative](wdshapesizerelative-enumeration-word.md)** enumeration), this property should be ignored because the shape does not use percent sizing. The height is solely determined by the **[Height](shaperange-height-property-word.md)** property.


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

