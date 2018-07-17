---
title: TextEffectFormat.RotatedChars Property (Publisher)
keywords: vbapb10.chm3735817
f1_keywords:
- vbapb10.chm3735817
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.RotatedChars
ms.assetid: 47566497-7b78-65dc-48d9-26b2e4245d31
ms.date: 06/08/2017
---


# TextEffectFormat.RotatedChars Property (Publisher)

 **msoTrue** if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. **msoFalse** if characters in the specified WordArt retain their original orientation relative to the bounding shape. Read/write.


## Syntax

 _expression_. **RotatedChars**

 _expression_A variable that represents a  **TextEffectFormat** object.


### Return Value

MsoTriState


## Remarks

If the WordArt has horizontal text, setting the  **RotatedChars** property to **True** rotates the characters 90 degrees counterclockwise. If the WordArt has vertical text, setting the **RotatedChars** property to **False** rotates the characters 90 degrees clockwise. Use the **[ToggleVerticalText](texteffectformat-toggleverticaltext-method-publisher.md)** method to switch between horizontal and vertical text flow.

The  **[Flip](shape-flip-method-publisher.md)** method and  **[Rotation](shape-rotation-property-publisher.md)** property of the  **[Shape](shape-object-publisher.md)** object and the  **RotatedChars** property and **ToggleVerticalText** method of the **[TextEffectFormat](texteffectformat-object-publisher.md)** object all affect the character orientation and direction of text flow in a **Shape** object that represents WordArt. You may have to experiment to find out how to combine the effects of these properties and methods to get the result you want.


## Example

This example adds WordArt that contains the text "Test" to the active publication and rotates the characters 90 degrees counterclockwise.


```vb
Sub CreateFormatWordArt() 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextEffect(PresetTextEffect:=msoTextEffect1, _ 
 Text:="Test", FontName:="Arial Black", FontSize:=36, _ 
 FontBold:=False, FontItalic:=False, Left:=10, Top:=10) 
 .TextEffect.RotatedChars = msoTrue 
 End With 
End Sub
```


