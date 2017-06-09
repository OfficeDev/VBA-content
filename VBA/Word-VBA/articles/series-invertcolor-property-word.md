---
title: Series.InvertColor Property (Word)
keywords: vbawd10.chm123734852
f1_keywords:
- vbawd10.chm123734852
ms.prod: word
api_name:
- Word.Series.InvertColor
ms.assetid: 50f248c7-5136-e4ea-c77c-9c0020275f07
ms.date: 06/08/2017
---


# Series.InvertColor Property (Word)

Returns or sets the fill color for negative data points in a series. Read/write.


## Syntax

 _expression_ . **InvertColor**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


### Return Value

Integer


## Remarks

The  **InvertColor** property enables you to set the fill color for negative data points as a specific numeric, hexadecimal, octal, or RGB color value. To set the value as an RBG value, use the Visual Basic[RGB](http://msdn.microsoft.com/library/5e9956de-ba18-56cd-0556-715774055cf4%28Office.15%29.aspx) function. Instead of using the **InvertColor** property, you can use the[InvertColorIndex](series-invertcolorindex-property-word.md) property, which uses a simplier set of integer values from the current color palette. For the **InvertColor** property to have an effect, the[InvertIfNegative](series-invertifnegative-property-word.md) property of the **Series** object must also be set to **True** .


## See also


#### Concepts


[Series Object](series-object-word.md)

