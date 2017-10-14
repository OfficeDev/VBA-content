---
title: Shape.TopRelative Property (Word)
keywords: vbawd10.chm161480905
f1_keywords:
- vbawd10.chm161480905
ms.prod: word
api_name:
- Word.Shape.TopRelative
ms.assetid: 5ae905f1-2e86-2aab-fe43-3be81f61847c
ms.date: 06/08/2017
---


# Shape.TopRelative Property (Word)

Returns or sets a  **Single** that represents the relative top position of a shape. Read/write.


## Syntax

 _expression_ . **TopRelative**

 _expression_ An expression that returns a **[Shape](shape-object-word.md)** object.


## Remarks

Use this property with the  **[RelativeHorizontalPosition](shape-relativehorizontalposition-property-word.md)** property. When set to **wdShapePositionRelativeNone** (-999999) (see the **[WdShapePositionRelative](wdshapepositionrelative-enumeration-word.md)** enumeration), this property should be ignored because the shape does not use percent positioning. The vertical position is solely determined by the **[Top](shape-top-property-word.md)** property.


## See also


#### Concepts


[Shape Object](shape-object-word.md)

