---
title: TextEffectFormat.NormalizedHeight Property (Word)
keywords: vbawd10.chm164561002
f1_keywords:
- vbawd10.chm164561002
ms.prod: word
api_name:
- Word.TextEffectFormat.NormalizedHeight
ms.assetid: 7410b830-3b1c-dc32-2ab8-c17a5a743c05
ms.date: 06/08/2017
---


# TextEffectFormat.NormalizedHeight Property (Word)

 **MsoTrue** if all characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write **MsoTriState** .


## Syntax

 _expression_ . **NormalizedHeight**

 _expression_ Required. A variable that represents a **[TextEffectFormat](texteffectformat-object-word.md)** object.


## Example

This example adds WordArt that contains the text "Test Effect" to myDocument and gives the new WordArt the name "texteff1." The code then makes all characters in the shape named "texteff1" the same height.


```vb
Set myDocument = ActiveDocument 
myDocument.Shapes.AddTextEffect(PresetTextEffect:=msoTextEffect1, _ 
 Text:="Test Effect", FontName:="Courier New", _ 
 FontSize:=44, FontBold:=True, _ 
 FontItalic:=False, Left:=10, Top:=10).Name = "texteff1" 
myDocument.Shapes("texteff1").TextEffect.NormalizedHeight = msoTrue
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-word.md)

