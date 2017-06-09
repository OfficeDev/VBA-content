---
title: Series.Formula Property (Word)
keywords: vbawd10.chm123732229
f1_keywords:
- vbawd10.chm123732229
ms.prod: word
api_name:
- Word.Series.Formula
ms.assetid: 744473dd-b7f2-6f70-a285-ddc95ef5221f
ms.date: 06/08/2017
---


# Series.Formula Property (Word)

Returns or sets the object's formula in A1-style notation and in the language of the macro. Read/write  **String** .


## Syntax

 _expression_ . **Formula**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

This property is not available for OLAP data sources.

If the object contains a constant, this property returns the constant. If the object is empty, this property returns an empty string. If the object contains a formula, this property returns the formula as a string in the same format that would be displayed in the formula bar (including the equal sign).

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.


## See also


#### Concepts


[Series Object](series-object-word.md)

