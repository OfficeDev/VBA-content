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
|<strong>ppLayoutBlank</strong>|
|
<strong>ppLayoutChart</strong>|
|
<strong>ppLayoutChartAndText</strong>|
|
<strong>ppLayoutClipartAndText</strong>|
|
<strong>ppLayoutClipArtAndVerticalText</strong>|
|
<strong>ppLayoutFourObjects</strong>|
|
<strong>ppLayoutLargeObject</strong>|
|
<strong>ppLayoutMediaClipAndText</strong>|
|
<strong>ppLayoutMixed</strong>|
|
<strong>ppLayoutObject</strong>|
|
<strong>ppLayoutObjectAndText</strong>|
|
<strong>ppLayoutObjectOverText</strong>|
|
<strong>ppLayoutOrgchart</strong>|
|
<strong>ppLayoutTable</strong>|
|
<strong>ppLayoutText</strong>|
|
<strong>ppLayoutTextAndChart</strong>|
|
<strong>ppLayoutTextAndClipart</strong>|
|
<strong>ppLayoutTextAndMediaClip</strong>|
|
<strong>ppLayoutTextAndObject</strong>|
|
<strong>ppLayoutTextAndTwoObjects</strong>|
|
<strong>ppLayoutTextOverObject</strong>|
|
<strong>ppLayoutTitle</strong>|
|
<strong>ppLayoutTitleOnly</strong>|
|
<strong>ppLayoutTwoColumnText</strong>|
|
<strong>ppLayoutTwoObjectsAndText</strong>|
|
<strong>ppLayoutTwoObjectsOverText</strong>|
|
<strong>ppLayoutVerticalText</strong>|
|
<strong>ppLayoutVerticalTitleAndText</strong>|
|
<strong>ppLayoutVerticalTitleAndTextOverChart</strong>|

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

