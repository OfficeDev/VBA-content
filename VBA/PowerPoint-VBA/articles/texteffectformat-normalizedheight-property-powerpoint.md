---
title: TextEffectFormat.NormalizedHeight Property (PowerPoint)
keywords: vbapp10.chm556009
f1_keywords:
- vbapp10.chm556009
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.NormalizedHeight
ms.assetid: 89b1799f-c037-5a37-caad-3344292df6e8
ms.date: 06/08/2017
---


# TextEffectFormat.NormalizedHeight Property (PowerPoint)

Determines whether the characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write.


## Syntax

 _expression_. **NormalizedHeight**

 _expression_ A variable that represents a **TextEffectFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **NormalizedHeight** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| All characters (both uppercase and lowercase) in the specified WordArt are not the same height.|
|**msoTrue**| All characters (both uppercase and lowercase) in the specified WordArt are the same height.|

## Example

This example adds WordArt that contains the text "Test Effect" to  `myDocument` and gives the new WordArt the name "texteff1." The code then makes all characters in the shape named "texteff1" the same height.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddTextEffect(PresetTextEffect:=msoTextEffect1, _
    Text:="Test Effect", FontName:="Courier New", _
    FontSize:=44, FontBold:=True, _
    FontItalic:=False, Left:=10, Top:=10)_
    .Name = "texteff1"

myDocument.Shapes("texteff1").TextEffect.NormalizedHeight = msoTrue
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-powerpoint.md)

