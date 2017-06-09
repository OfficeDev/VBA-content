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
|**msoTextEffectShapeArchDownCurve**|
|**msoTextEffectShapeArchDownPour**|
|**msoTextEffectShapeArchUpCurve**|
|**msoTextEffectShapeArchUpPour**|
|**msoTextEffectShapeButtonCurve**|
|**msoTextEffectShapeButtonPour**|
|**msoTextEffectShapeCanDown**|
|**msoTextEffectShapeCanUp**|
|**msoTextEffectShapeCascadeDown**|
|**msoTextEffectShapeCascadeUp**|
|**msoTextEffectShapeChevronDown**|
|**msoTextEffectShapeChevronUp**|
|**msoTextEffectShapeCircleCurve**|
|**msoTextEffectShapeCirclePour**|
|**msoTextEffectShapeCurveDown**|
|**msoTextEffectShapeCurveUp**|
|**msoTextEffectShapeDeflate**|
|**msoTextEffectShapeDeflateBottom**|
|**msoTextEffectShapeDeflateInflate**|
|**msoTextEffectShapeDeflateInflateDeflate**|
|**msoTextEffectShapeDeflateTop**|
|**msoTextEffectShapeDoubleWave2**|
|**msoTextEffectShapeFadeDown**|
|**msoTextEffectShapeFadeLeft**|
|**msoTextEffectShapeFadeRight**|
|**msoTextEffectShapeFadeUp**|
|**msoTextEffectShapeInflate**|
|**msoTextEffectShapeInflateBottom**|
|**msoTextEffectShapeInflateTop**|
|**msoTextEffectShapeMixed**|
|**msoTextEffectShapePlainText**|
|**msoTextEffectShapeRingInside**|
|**msoTextEffectShapeRingOutside**|
|**msoTextEffectShapeSlantDown**|
|**msoTextEffectShapeSlantUp**|
|**msoTextEffectShapeStop**|
|**msoTextEffectShapeTriangleDown**|
|**msoTextEffectShapeTriangleUp**|
|**msoTextEffectShapeWave1**|
|**msoTextEffectShapeWave2**|
|**msoTextEffectShapeDoubleWave1**|

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

