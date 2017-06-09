---
title: TextEffectFormat.ToggleVerticalText Method (Excel)
keywords: vbaxl10.chm118020
f1_keywords:
- vbaxl10.chm118020
ms.prod: excel
api_name:
- Excel.TextEffectFormat.ToggleVerticalText
ms.assetid: 9b4312b8-1642-9a49-6395-b49b129f44f2
ms.date: 06/08/2017
---


# TextEffectFormat.ToggleVerticalText Method (Excel)

Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.


## Syntax

 _expression_ . **ToggleVerticalText**

 _expression_ A variable that represents a **TextEffectFormat** object.


## Remarks

Using the  **ToggleVerticalText** method swaps the values of the **[Width](shape-width-property-excel.md)** and **[Height](shape-height-property-excel.md)** properties of the **[Shape](shape-object-excel.md)** object that represents the WordArt and leaves the **[Left](shape-left-property-excel.md)** and **[Top](shape-top-property-excel.md)** properties unchanged.

The  **[Flip](shape-flip-method-excel.md)** method and **[Rotation](shape-rotation-property-excel.md)** property of the **Shape** object and the **[RotatedChars](texteffectformat-rotatedchars-property-excel.md)** property and **[ToggleVerticalText](texteffectformat-toggleverticaltext-method-excel.md)** method of the **[TextEffectFormat](texteffectformat-object-excel.md)** object all affect the character orientation and the direction of text flow in a **Shape** object that represents WordArt. You may have to experiment to find out how to combine the effects of these properties and methods to get the result you want.


## Example

This example adds WordArt that contains the text "Test" to  `myDocument` and switches from horizontal text flow (the default for the specified WordArt style, **msoTextEffect1** ) to vertical text flow.


```vb
Set myDocument = Worksheets(1) 
Set newWordArt = myDocument.Shapes.AddTextEffect( _ 
 PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
 FontName:="Arial Black", FontSize:=36, _ 
 FontBold:=False, FontItalic:=False, Left:=100, _ 
 Top:=100) 
newWordArt.TextEffect.ToggleVerticalText
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-excel.md)

