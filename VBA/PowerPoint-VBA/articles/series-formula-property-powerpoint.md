---
title: Series.Formula Property (PowerPoint)
keywords: vbapp10.chm65797
f1_keywords:
- vbapp10.chm65797
ms.prod: powerpoint
api_name:
- PowerPoint.Series.Formula
ms.assetid: 04d62f5d-e63d-1643-a6cd-eae0c37b73cf
ms.date: 06/08/2017
---


# Series.Formula Property (PowerPoint)

Returns or sets the object's formula in A1-style notation and in the language of the macro. Read/write  **String**.


## Syntax

 _expression_. **Formula**

 _expression_ A variable that represents a **[Series](series-object-powerpoint.md)** object.


## Remarks

This property is not available for OLAP data sources.

If the object contains a constant, this property returns the constant. If the object is empty, this property returns an empty string. If the object contains a formula, this property returns the formula as a string in the same format that would be displayed in the formula bar (including the equal sign).

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

