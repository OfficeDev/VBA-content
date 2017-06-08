---
title: ShapeRange.WidthRelative Property (Word)
keywords: vbawd10.chm162857162
f1_keywords:
- vbawd10.chm162857162
ms.prod: word
api_name:
- Word.ShapeRange.WidthRelative
ms.assetid: 907626b9-80e2-ea63-d6a6-27295ef1e2c4
ms.date: 06/08/2017
---


# ShapeRange.WidthRelative Property (Word)

Returns or sets a  **Single** that represents the relative width of a range of shapes. Read/write.


## Syntax

 _expression_ . **WidthRelative**

 _expression_ An expression that returns a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

Use this property with the  **[RelativeHorizontalSize](shaperange-relativehorizontalsize-property-word.md)** property. When set to **wdShapeSizeRelativeNone** (-999999) (see the **[WdShapeSizeRelative](wdshapesizerelative-enumeration-word.md)** enumeration), this property should be ignored because the shape does not use percent sizing. The width is solely determined by the **[Width](shaperange-width-property-word.md)** property.


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

