---
title: TextEffectFormat.FontSize Property (Word)
keywords: vbawd10.chm164561000
f1_keywords:
- vbawd10.chm164561000
ms.prod: word
api_name:
- Word.TextEffectFormat.FontSize
ms.assetid: 14538296-38d0-0545-0681-e6a7714dcaf4
ms.date: 06/08/2017
---


# TextEffectFormat.FontSize Property (Word)

Returns or sets the font size for the specified WordArt, in points. Read/write  **Single** .


## Syntax

 _expression_ . **FontSize**

 _expression_ A variable that represents a **[TextEffectFormat](texteffectformat-object-word.md)** object.


## Example

This example sets the font size to 16 points for the shape named "WordArt 2" in the active document.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
docActive.Shapes("WordArt 2").TextEffect.FontSize = 16
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-word.md)

