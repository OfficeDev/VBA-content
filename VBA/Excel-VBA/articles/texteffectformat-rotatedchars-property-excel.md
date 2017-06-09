---
title: TextEffectFormat.RotatedChars Property (Excel)
keywords: vbaxl10.chm118011
f1_keywords:
- vbaxl10.chm118011
ms.prod: excel
api_name:
- Excel.TextEffectFormat.RotatedChars
ms.assetid: 708f076d-82e7-f7f3-a2df-3f4a4d628092
ms.date: 06/08/2017
---


# TextEffectFormat.RotatedChars Property (Excel)

 **True** if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. **False** if characters in the specified WordArt retain their original orientation relative to the bounding shape. Read/write **MsoTriState** .


## Syntax

 _expression_ . **RotatedChars**

 _expression_ A variable that represents a **TextEffectFormat** object.


## Remarks



| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse** Characters in the specified WordArt retain their original orientation relative to the bounding shape.|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** Characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape.|
If the WordArt has horizontal text, setting the  **RotatedChars** property to **msoTrue** rotates the characters 90 degrees counterclockwise. If the WordArt has vertical text, setting the **RotatedChars** property to **msoFalse** rotates the characters 90 degrees clockwise. Use the **ToggleVerticalText** method to switch between horizontal and vertical text flow.

The  **[Flip](shape-flip-method-excel.md)** method and **[Rotation](shape-rotation-property-excel.md)** property of the **[Shape](shape-object-excel.md)** object and the **RotatedChars** property and **[ToggleVerticalText](texteffectformat-toggleverticaltext-method-excel.md)** method of the **TextEffectFormat** object all affect the character orientation and direction of text flow in a **Shape** object that represents WordArt. You may have to experiment to find out how to combine the effects of these properties and methods to get the result you want.


## Example

This example adds WordArt that contains the text "Test" to  `myDocument` and rotates the characters 90 degrees counterclockwise.


```vb
Set myDocument = Worksheets(1) 
Set newWordArt = myDocument.Shapes.AddTextEffect( _ 
 PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
 FontName:="Arial Black", FontSize:=36, _ 
 FontBold:=False, FontItalic:=False, Left:=10, _ 
 Top:=10) 
newWordArt.TextEffect.RotatedChars = msoTrue
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-excel.md)

