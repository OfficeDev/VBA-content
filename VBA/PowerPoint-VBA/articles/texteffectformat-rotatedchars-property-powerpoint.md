---
title: TextEffectFormat.RotatedChars Property (PowerPoint)
keywords: vbapp10.chm556012
f1_keywords:
- vbapp10.chm556012
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.RotatedChars
ms.assetid: ae12e31d-d86b-73d2-ab92-a2d6bc8a2036
ms.date: 06/08/2017
---


# TextEffectFormat.RotatedChars Property (PowerPoint)

Determines whether characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. Read/write.


## Syntax

 _expression_. **RotatedChars**

 _expression_ A variable that represents a **TextEffectFormat** object.


### Return Value

MsoTriState


## Remarks

If the WordArt has horizontal text, setting the  **RotatedChars** property to **msoTrue** rotates the characters 90 degrees counterclockwise. If the WordArt has vertical text, setting the **RotatedChars** property to **msoFalse** rotates the characters 90 degrees clockwise. Use the **ToggleVerticalText** method to switch between horizontal and vertical text flow.

The  **[Flip](shape-flip-method-powerpoint.md)** method and **[Rotation](shape-rotation-property-powerpoint.md)** property of the **[Shape](shape-object-powerpoint.md)** object and the **RotatedChars** property and **[ToggleVerticalText](texteffectformat-toggleverticaltext-method-powerpoint.md)** method of the **TextEffectFormat** object all affect the character orientation and direction of text flow in a **Shape** object that represents WordArt. You may have to experiment to find out how to combine the effects of these properties and methods to get the result you want.

The value of the  **RotatedChars** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Characters in the specified WordArt retain their original orientation relative to the bounding shape.|
|**msoTrue**| Characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape.|

## Example

This example adds WordArt that contains the text "Test" to  `myDocument` and rotates the characters 90 degrees counterclockwise.


```vb
Set myDocument = ActivePresentation.Slides(1)

Set newWordArt = myDocument.Shapes.AddTextEffect _
    (PresetTextEffect:=msoTextEffect1, Text:="Test", _
    FontName:="Arial Black", FontSize:=36, _
    FontBold:=msoFalse, FontItalic:=msoFalse, Left:=10, Top:=10)

newWordArt.TextEffect.RotatedChars = msoTrue
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-powerpoint.md)

