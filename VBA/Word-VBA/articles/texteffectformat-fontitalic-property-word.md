---
title: TextEffectFormat.FontItalic Property (Word)
keywords: vbawd10.chm164560998
f1_keywords:
- vbawd10.chm164560998
ms.prod: word
api_name:
- Word.TextEffectFormat.FontItalic
ms.assetid: a5fa97ea-c01d-8742-9e9e-20a8148a3326
ms.date: 06/08/2017
---


# TextEffectFormat.FontItalic Property (Word)

Italicizes WordArt text. Read/write  **MsoTriState** .


## Syntax

 _expression_ . **FontItalic**

 _expression_ Required. A variable that represents a **[TextEffectFormat](texteffectformat-object-word.md)** object.


## Example

This example sets the font to italic for the shape named "WordArt 4" in the active document.


```vb
Sub ItalicizeWordArt() 
 ActiveDocument.Shapes("WordArt 4") _ 
 .TextEffect.FontItalic = msoTrue 
End Sub
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-word.md)

