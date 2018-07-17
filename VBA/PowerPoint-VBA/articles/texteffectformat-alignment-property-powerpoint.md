---
title: TextEffectFormat.Alignment Property (PowerPoint)
keywords: vbapp10.chm556003
f1_keywords:
- vbapp10.chm556003
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.Alignment
ms.assetid: 42b92de7-2dc1-ee1b-1c58-682cfba2aa19
ms.date: 06/08/2017
---


# TextEffectFormat.Alignment Property (PowerPoint)

Returns or sets the alignment for the specified WordArt. Read/write.


## Syntax

 _expression_. **Alignment**

 _expression_ A variable that represents a **TextEffectFormat** object.


## Remarks

The value of the  **Alignment** property can be one of these **MsoTextEffectAlignment** constants.


||
|:-----|
|**msoTextEffectAlignmentCentered**|
|**msoTextEffectAlignmentLeft**|
|**msoTextEffectAlignmentMixed**|
|**msoTextEffectAlignmentRight**|
|**msoTextEffectAlignmentStretchJustify**|
|**msoTextEffectAlignmentWordJustify**|
|**msoTextEffectAlignmentLetterJustify**|

## Example

This example adds a WordArt object to the first slide in the active presentation and then right-aligns the WordArt.


```vb
Set mySh = Application.ActivePresentation.Slides(1).Shapes

Set myTE = mySh.AddTextEffect(PresetTextEffect:=msoTextEffect1, _
    Text:="Test Text", FontName:="Palatino", FontSize:=54, _
    FontBold:=True, FontItalic:=False, Left:=100, Top:=50)

myTE.TextEffect.Alignment = msoTextEffectAlignmentRight
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-powerpoint.md)

