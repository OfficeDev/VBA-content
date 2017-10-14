---
title: TextEffectFormat.RotatedChars Property (Word)
keywords: vbawd10.chm164561005
f1_keywords:
- vbawd10.chm164561005
ms.prod: word
api_name:
- Word.TextEffectFormat.RotatedChars
ms.assetid: 4f5c9f84-0c86-1558-ac64-ca8d53e3683d
ms.date: 06/08/2017
---


# TextEffectFormat.RotatedChars Property (Word)

 **MsoTrue** if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. **MsoFalse** if characters in the specified WordArt retain their original orientation relative to the bounding shape. Read/write **MsoTriState** .


## Syntax

 _expression_ . **RotatedChars**

 _expression_ Required. A variable that represents a **[TextEffectFormat](texteffectformat-object-word.md)** object.


## Remarks

If the WordArt has horizontal text, setting the  **RotatedChars** property to **True** rotates the characters 90 degrees counterclockwise. If the WordArt has vertical text, setting the **RotatedChars** property to **False** rotates the characters 90 degrees clockwise. Use the **ToggleVerticalText** method to switch between horizontal and vertical text flow.

The  **Flip** method and **Rotation** property of the **Shape** object and the **RotatedChars** property and **ToggleVerticalText** method of the **TextEffectFormat** object all affect the character orientation and direction of text flow in a **[Shape](shape-object-word.md)** object that represents WordArt. You may have to experiment to find out how to combine the effects of these properties and methods to get the result you want.


## Example

This example adds WordArt that contains the text "Test" to myDocument and rotates the characters 90 degrees counterclockwise.


```vb
Set myDocument = ActiveDocument 
Set newWordArt = _ 
 myDocument.Shapes.AddTextEffect( _ 
 PresetTextEffect:=msoTextEffect1, _ 
 Text:="Test", _ 
 FontName:="Arial Black", FontSize:=36, _ 
 FontBold:=False, FontItalic:=False, Left:=10, Top:=10) 
newWordArt.TextEffect.RotatedChars = True
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-word.md)

