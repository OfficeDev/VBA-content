---
title: Slide.Layout Property (PowerPoint)
keywords: vbapp10.chm531014
f1_keywords:
- vbapp10.chm531014
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Layout
ms.assetid: 681819b8-327e-fb6f-e9d2-0f8feb48ec36
ms.date: 06/08/2017
---


# Slide.Layout Property (PowerPoint)

Returns or sets a  **PpSlideLayout** constant that represents the slide layout. Read/write.


## Syntax

 _expression_. **Layout**

 _expression_ A variable that represents a **Slide** object.


## Remarks

The value of the  **Layout** property can be one of these **PpSlideLayout** constants.


||
|:-----|
|**ppLayoutBlank**|
|**ppLayoutChart**|
|**ppLayoutChartAndText**|
|**ppLayoutClipartAndText**|
|**ppLayoutClipArtAndVerticalText**|
|**ppLayoutFourObjects**|
|**ppLayoutLargeObject**|
|**ppLayoutMediaClipAndText**|
|**ppLayoutMixed**|
|**ppLayoutObject**|
|**ppLayoutObjectAndText**|
|**ppLayoutObjectOverText**|
|**ppLayoutOrgchart**|
|**ppLayoutTable**|
|**ppLayoutText**|
|**ppLayoutTextAndChart**|
|**ppLayoutTextAndClipart**|
|**ppLayoutTextAndMediaClip**|
|**ppLayoutTextAndObject**|
|**ppLayoutTextAndTwoObjects**|
|**ppLayoutTextOverObject**|
|**ppLayoutTitle**|
|**ppLayoutTitleOnly**|
|**ppLayoutTwoColumnText**|
|**ppLayoutTwoObjectsAndText**|
|**ppLayoutTwoObjectsOverText**|
|**ppLayoutVerticalText**|
|**ppLayoutVerticalTitleAndText**|
|**ppLayoutVerticalTitleAndTextOverChart**|

## Example

This example changes the layout of slide one in the active presentation to include a title and subtitle if it initially has only a title.


```vb
With ActivePresentation.Slides(1)

    If .Layout = ppLayoutTitleOnly Then

        .Layout = ppLayoutTitle

    End If

End With
```


## See also


#### Concepts


[Slide Object](slide-object-powerpoint.md)

