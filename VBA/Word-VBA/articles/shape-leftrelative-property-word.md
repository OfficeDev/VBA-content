---
title: Shape.LeftRelative Property (Word)
keywords: vbawd10.chm161480904
f1_keywords:
- vbawd10.chm161480904
ms.prod: word
api_name:
- Word.Shape.LeftRelative
ms.assetid: a4fd7e18-9e04-8ea9-58d1-e2e816079ac3
ms.date: 06/08/2017
---


# Shape.LeftRelative Property (Word)

Returns or sets a  **Single** that represents the relative left position of a shape. Read/write.


## Syntax

 _expression_ . **LeftRelative**

 _expression_ An expression that returns a **[Shape](shape-object-word.md)** object.


## Remarks

Use this property with the  **[RelativeHorizontalPosition](shape-relativehorizontalposition-property-word.md)** property. When set to **wdShapePositionRelativeNone** (-999999) (see the **[WdShapePositionRelative](wdshapepositionrelative-enumeration-word.md)** enumeration), this property should be ignored because the shape does not use percent positioning. The horizontal position is solely determined by the **[Left](shape-left-property-word.md)** property.


## See also


#### Concepts


[Shape Object](shape-object-word.md)

