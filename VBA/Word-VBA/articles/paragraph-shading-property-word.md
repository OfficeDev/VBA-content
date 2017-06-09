---
title: Paragraph.Shading Property (Word)
keywords: vbawd10.chm156696692
f1_keywords:
- vbawd10.chm156696692
ms.prod: word
api_name:
- Word.Paragraph.Shading
ms.assetid: 870ddeb5-e2fe-ff77-baac-7270a307be7c
ms.date: 06/08/2017
---


# Paragraph.Shading Property (Word)

Returns a  **[Shading](shading-object-word.md)** object that refers to the shading formatting for the specified paragraph.


## Syntax

 _expression_ . **Shading**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example applies yellow shading to the first paragraph in the selection.


```vb
With Selection.Paragraphs(1).Shading 
 .Texture = wdTexture12Pt5Percent 
 .BackgroundPatternColorIndex = wdYellow 
 .ForegroundPatternColorIndex = wdBlack 
End With
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

