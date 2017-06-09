---
title: TextEffectFormat.ToggleVerticalText Method (Publisher)
keywords: vbapb10.chm3735568
f1_keywords:
- vbapb10.chm3735568
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.ToggleVerticalText
ms.assetid: 627ddbcc-5951-70c6-4e54-de0e9a4bebec
ms.date: 06/08/2017
---


# TextEffectFormat.ToggleVerticalText Method (Publisher)

Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.


## Syntax

 _expression_. **ToggleVerticalText**

 _expression_A variable that represents a  **TextEffectFormat** object.


## Remarks

Using the  **ToggleVerticalText** method swaps the values of the **[Left](shape-left-property-publisher.md)** and **[Top](shape-top-property-publisher.md)** properties of the **[Shape](shape-object-publisher.md)** object that represents the WordArt and leaves the  **[Width](shape-width-property-publisher.md)** and **[Height](shape-height-property-publisher.md)** properties unchanged.

The  **[Flip](shape-flip-method-publisher.md)** method and  **[Rotation](shape-rotation-property-publisher.md)** property of the  **[Shape](shape-object-publisher.md)** object and the  **[RotatedChars](texteffectformat-rotatedchars-property-publisher.md)** property and  **ToggleVerticalText** method of the **[TextEffectFormat](texteffectformat-object-publisher.md)** object all affect the character orientation and the direction of text flow in a **Shape** object that represents WordArt. You may have to experiment to find out how to combine the effects of these properties and methods to get the result you want.


## Example

This example adds WordArt that contains the text "Test" to the active publication, and switches from horizontal text flow (the default for the specified WordArt style,  **msoTextEffect1**) to vertical text flow.


```vb
Dim shpTextEffect As Shape 
 
Set shpTextEffect = ActiveDocument.Pages(1).Shapes.AddTextEffect _ 
 (PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
 FontName:="Arial Black", FontSize:=36, _ 
 FontBold:=False, FontItalic:=False, Left:=100, Top:=100) 
 
shpTextEffect.TextEffect.ToggleVerticalText
```


