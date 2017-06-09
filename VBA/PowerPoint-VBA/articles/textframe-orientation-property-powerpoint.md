---
title: TextFrame.Orientation Property (PowerPoint)
keywords: vbapp10.chm558006
f1_keywords:
- vbapp10.chm558006
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.Orientation
ms.assetid: ce6a9578-3cbd-9b73-e374-c43fa4748054
ms.date: 06/08/2017
---


# TextFrame.Orientation Property (PowerPoint)

Returns or sets text orientation. Read/write.


## Syntax

 _expression_. **Orientation**

 _expression_ A variable that represents a **TextFrame** object.


### Return Value

MsoTextOrientation


## Remarks

Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.

The value of the  **Orientation** property can be one of these **MsoTextOrientation** constants.


||
|:-----|
|**msoTextOrientationDownward**|
|**msoTextOrientationHorizontal**|
|**msoTextOrientationHorizontalRotatedFarEast**|
|**msoTextOrientationMixed**|
|**msoTextOrientationUpward**|
|**msoTextOrientationVertical**|
|**msoTextOrientationVerticalFarEast**|

## Example

This example orients the text horizontally within shape three on myDocument.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(3).TextFrame.Orientation = msoTextOrientationHorizontal


```


## See also


#### Concepts


[TextFrame Object](textframe-object-powerpoint.md)

