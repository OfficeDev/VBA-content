---
title: Paragraphs.Shading Property (Word)
keywords: vbawd10.chm156762228
f1_keywords:
- vbawd10.chm156762228
ms.prod: word
api_name:
- Word.Paragraphs.Shading
ms.assetid: b732c59d-d861-00d8-fd00-6940449480a1
ms.date: 06/08/2017
---


# Paragraphs.Shading Property (Word)

Returns a  **Shading** object that refers to the shading formatting for the specified paragraphs.


## Syntax

 _expression_ . **Shading**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example applies yellow shading to the all paragraphs in the selection.


```vb
With Selection.Paragraphs.Shading 
 .Texture = wdTexture12Pt5Percent 
 .BackgroundPatternColorIndex = wdYellow 
 .ForegroundPatternColorIndex = wdBlack 
End With
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

