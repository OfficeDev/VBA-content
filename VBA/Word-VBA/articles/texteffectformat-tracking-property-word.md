---
title: TextEffectFormat.Tracking Property (Word)
keywords: vbawd10.chm164561007
f1_keywords:
- vbawd10.chm164561007
ms.prod: word
api_name:
- Word.TextEffectFormat.Tracking
ms.assetid: 40e1ac58-b292-ac12-6e82-a93f87013d6d
ms.date: 06/08/2017
---


# TextEffectFormat.Tracking Property (Word)

Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt in relation to the width of the character. Read/write  **Single** .


## Syntax

 _expression_ . **Tracking**

 _expression_ Required. A variable that represents a **[TextEffectFormat](texteffectformat-object-word.md)** object.


## Remarks

This property can be a value from 0 (zero) through 5. (Large values for this property specify ample space between characters; values less than 1 can produce character overlap.) The following table gives the values of the  **Tracking** property that correspond to the settings available in the user interface.



|**User interface setting**|**Equivalent Tracking property value**|
|:-----|:-----|
|Very Tight|0.8|
|Tight|0.9|
|Normal|1.0|
|Loose|1.2|
|Very Loose|1.5|

## Example

This example adds WordArt that contains the text "Test" to the active document and specifies that the characters be very tightly spaced.


```vb
Set newWordArt = ActiveDocument.Shapes.AddTextEffect( _ 
 PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
 FontName:="Arial Black", FontSize:=36, FontBold:=False, _ 
 FontItalic:=False, Left:=100, Top:=100) 
newWordArt.TextEffect.Tracking = 0.8
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-word.md)

