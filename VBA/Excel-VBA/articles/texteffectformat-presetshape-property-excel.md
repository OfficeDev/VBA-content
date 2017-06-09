---
title: TextEffectFormat.PresetShape Property (Excel)
keywords: vbaxl10.chm118009
f1_keywords:
- vbaxl10.chm118009
ms.prod: excel
api_name:
- Excel.TextEffectFormat.PresetShape
ms.assetid: d85bcdf6-0ad4-4a3d-ed3a-40a96a4b63a0
ms.date: 06/08/2017
---


# TextEffectFormat.PresetShape Property (Excel)

Returns or sets the shape of the specified WordArt. Read/write  **MsoPresetTextEffectShape** .


## Syntax

 _expression_ . **PresetShape**

 _expression_ A variable that represents a **TextEffectFormat** object.


## Remarks



| **MsoPresetTextEffectShape** can be one of these **MsoPresetTextEffectShape** constants.|
| **msoTextEffectShapeArchDownCurve**|
| **msoTextEffectShapeArchDownPour**|
| **msoTextEffectShapeArchUpCurve**|
| **msoTextEffectShapeArchUpPour**|
| **msoTextEffectShapeButtonCurve**|
| **msoTextEffectShapeButtonPour**|
| **msoTextEffectShapeCanDown**|
| **msoTextEffectShapeCanUp**|
| **msoTextEffectShapeCascadeDown**|
| **msoTextEffectShapeCascadeUp**|
| **msoTextEffectShapeChevronDown**|
| **msoTextEffectShapeChevronUp**|
| **msoTextEffectShapeCircleCurve**|
| **msoTextEffectShapeCirclePour**|
| **msoTextEffectShapeCurveDown**|
| **msoTextEffectShapeCurveUp**|
| **msoTextEffectShapeDeflate**|
| **msoTextEffectShapeDeflateBottom**|
| **msoTextEffectShapeDeflateInflateDeflate**|
| **msoTextEffectShapeDoubleWave1**|
| **msoTextEffectShapeFadeDown**|
| **msoTextEffectShapeFadeRight**|
| **msoTextEffectShapeInflate**|
| **msoTextEffectShapeInflateTop**|
| **msoTextEffectShapePlainText**|
| **msoTextEffectShapeRingOutside**|
| **msoTextEffectShapeSlantUp**|
| **msoTextEffectShapeTriangleDown**|
| **msoTextEffectShapeWave1**|
| **msoTextEffectShapeDeflateInflate**|
| **msoTextEffectShapeDeflateTop**|
| **msoTextEffectShapeDoubleWave2**|
| **msoTextEffectShapeFadeLeft**|
| **msoTextEffectShapeFadeUp**|
| **msoTextEffectShapeInflateBottom**|
| **msoTextEffectShapeMixed**|
| **msoTextEffectShapeRingInside**|
| **msoTextEffectShapeSlantDown**|
| **msoTextEffectShapeStop**|
| **msoTextEffectShapeTriangleUp**|
| **msoTextEffectShapeWave2**|
Setting the  **[PresetTextEffect](texteffectformat-presettexteffect-property-excel.md)** property automatically sets the **PresetShape** property.


## Example

This example sets the shape of all WordArt on  `myDocument` to a chevron whose center points down.


```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
 If s.Type = msoTextEffect Then 
 s.TextEffect.PresetShape = msoTextEffectShapeChevronDown 
 End If 
Next
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-excel.md)

