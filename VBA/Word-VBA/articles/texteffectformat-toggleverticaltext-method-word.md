---
title: TextEffectFormat.ToggleVerticalText Method (Word)
keywords: vbawd10.chm164560906
f1_keywords:
- vbawd10.chm164560906
ms.prod: word
api_name:
- Word.TextEffectFormat.ToggleVerticalText
ms.assetid: 3d6fb851-e6f4-d8fc-a37a-80fb9455ca81
ms.date: 06/08/2017
---


# TextEffectFormat.ToggleVerticalText Method (Word)

Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.


## Syntax

 _expression_ . **ToggleVerticalText**

 _expression_ Required. A variable that represents a **[TextEffectFormat](texteffectformat-object-word.md)** object.


## Remarks

Using the  **ToggleVerticalText** method swaps the values of the **Width** and **Height** properties of the **[Shape](shape-object-word.md)** object that represents the WordArt and leaves the **Left** and **Top** properties unchanged.

The  **Flip** method and **Rotation** property of the **Shape** object and the **RotatedChars** property and **ToggleVerticalText** method of the **TextEffectFormat** object all affect the character orientation and the direction of text flow in a **[Shape](shape-object-word.md)** object that represents WordArt. You may have to experiment to find out how to combine the effects of these properties and methods to get the result you want.


## Example

This example adds WordArt that contains the text "Test" to the active document and switches from horizontal text flow (the default for the specified WordArt style,  **msoTextEffect1** ) to vertical text flow.


```vb
Dim newWordArt As Shape 
 
Set newWordArt = _ 
 ActiveDocument.Shapes.AddTextEffect( _ 
 PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
 FontName:="Arial Black", FontSize:=36, FontBold:=False, _ 
 FontItalic:=False, Left:=100, Top:=100) 
newWordArt.TextEffect.ToggleVerticalText
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-word.md)

