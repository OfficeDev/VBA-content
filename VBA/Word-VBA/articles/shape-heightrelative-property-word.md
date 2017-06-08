---
title: Shape.HeightRelative Property (Word)
keywords: vbawd10.chm161480907
f1_keywords:
- vbawd10.chm161480907
ms.prod: word
api_name:
- Word.Shape.HeightRelative
ms.assetid: 24a52ebf-1071-a2e4-8222-9b17d295e653
ms.date: 06/08/2017
---


# Shape.HeightRelative Property (Word)

Returns or sets a  **Single** that represents the percentage of the relative height of a shape. Read/write.


## Syntax

 _expression_ . **HeightRelative**

 _expression_ An expression that returns a **[Shape](shape-object-word.md)** object.


## Remarks

Use this property with the  **[RelativeVerticalSize](shape-relativeverticalsize-property-word.md)** property. When set to **wdShapeSizeRelativeNone** (-999999) (see the **[WdShapeSizeRelative](wdshapesizerelative-enumeration-word.md)** enumeration), this property should be ignored because the shape does not use percent sizing. The height is solely determined by the **[Height](shape-height-property-word.md)** property.


## See also


#### Concepts


[Shape Object](shape-object-word.md)

