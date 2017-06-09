---
title: TextFrame2.Orientation Property (PowerPoint)
keywords: vbapp10.chm678006
f1_keywords:
- vbapp10.chm678006
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.Orientation
ms.assetid: 713ce09e-575a-c1be-b60b-67884cb76673
ms.date: 06/08/2017
---


# TextFrame2.Orientation Property (PowerPoint)

 Returns or sets text orientation. Read/write.


## Syntax

 _expression_. **Orientation**

 _expression_ An expression that returns a **TextFrame2** object.


### Return Value

MsoTextOrientation


## Remarks

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

 **Note**  Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example shows how to orient the text horizontally in shape one on slide one in the active presentation.


```vb
Public Sub Orientation_Example()



    Dim pptSlide As Slide

    Set pptSlide = ActivePresentation.Slides(1)

    pptSlide.Shapes(1).TextFrame2.Orientation = msoTextOrientationHorizontal



End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-powerpoint.md)

