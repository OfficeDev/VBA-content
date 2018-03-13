---
title: TextEffectFormat.PresetShape Property (PowerPoint)
keywords: vbapp10.chm556010
f1_keywords:
- vbapp10.chm556010
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.PresetShape
ms.assetid: e4e43c4c-09fa-4f1d-a0de-26e0c7a872a0
ms.date: 06/08/2017
---


# TextEffectFormat.PresetShape Property (PowerPoint)

Returns or sets the shape of the specified WordArt. Read/write.


## Syntax

 _expression_. **PresetShape**

 _expression_ A variable that represents a **TextEffectFormat** object.


### Return Value

MsoPresetTextEffectShape


## Remarks

Setting the  **[PresetTextEffect](texteffectformat-presettexteffect-property-powerpoint.md)** property automatically sets the **PresetShape** property.

The value of the  **PresetShape** property can be one of these **MsoPresetTextEffectShape** constants.


||
|:-----|
|<strong>msoTextEffectShapeArchDownCurve</strong>|
|
<strong>msoTextEffectShapeArchDownPour</strong>|
|
<strong>msoTextEffectShapeArchUpCurve</strong>|
|
<strong>msoTextEffectShapeArchUpPour</strong>|
|
<strong>msoTextEffectShapeButtonCurve</strong>|
|
<strong>msoTextEffectShapeButtonPour</strong>|
|
<strong>msoTextEffectShapeCanDown</strong>|
|
<strong>msoTextEffectShapeCanUp</strong>|
|
<strong>msoTextEffectShapeCascadeDown</strong>|
|
<strong>msoTextEffectShapeCascadeUp</strong>|
|
<strong>msoTextEffectShapeChevronDown</strong>|
|
<strong>msoTextEffectShapeChevronUp</strong>|
|
<strong>msoTextEffectShapeCircleCurve</strong>|
|
<strong>msoTextEffectShapeCirclePour</strong>|
|
<strong>msoTextEffectShapeCurveDown</strong>|
|
<strong>msoTextEffectShapeCurveUp</strong>|
|
<strong>msoTextEffectShapeDeflate</strong>|
|
<strong>msoTextEffectShapeDeflateBottom</strong>|
|
<strong>msoTextEffectShapeDeflateInflate</strong>|
|
<strong>msoTextEffectShapeDeflateInflateDeflate</strong>|
|
<strong>msoTextEffectShapeDeflateTop</strong>|
|
<strong>msoTextEffectShapeDoubleWave2</strong>|
|
<strong>msoTextEffectShapeFadeDown</strong>|
|
<strong>msoTextEffectShapeFadeLeft</strong>|
|
<strong>msoTextEffectShapeFadeRight</strong>|
|
<strong>msoTextEffectShapeFadeUp</strong>|
|
<strong>msoTextEffectShapeInflate</strong>|
|
<strong>msoTextEffectShapeInflateBottom</strong>|
|
<strong>msoTextEffectShapeInflateTop</strong>|
|
<strong>msoTextEffectShapeMixed</strong>|
|
<strong>msoTextEffectShapePlainText</strong>|
|
<strong>msoTextEffectShapeRingInside</strong>|
|
<strong>msoTextEffectShapeRingOutside</strong>|
|
<strong>msoTextEffectShapeSlantDown</strong>|
|
<strong>msoTextEffectShapeSlantUp</strong>|
|
<strong>msoTextEffectShapeStop</strong>|
|
<strong>msoTextEffectShapeTriangleDown</strong>|
|
<strong>msoTextEffectShapeTriangleUp</strong>|
|
<strong>msoTextEffectShapeWave1</strong>|
|
<strong>msoTextEffectShapeWave2</strong>|
|
<strong>msoTextEffectShapeDoubleWave1</strong>|

## Example

This example sets the shape of all WordArt on  `myDocument` to a chevron whose center points down.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    If s.Type = msoTextEffect Then

        s.TextEffect.PresetShape = msoTextEffectShapeChevronDown

    End If

Next
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-powerpoint.md)

