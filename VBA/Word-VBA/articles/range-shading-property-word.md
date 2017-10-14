---
title: Range.Shading Property (Word)
keywords: vbawd10.chm157155389
f1_keywords:
- vbawd10.chm157155389
ms.prod: word
api_name:
- Word.Range.Shading
ms.assetid: 8e09cd74-a16e-6547-5ada-97322cf32b99
ms.date: 06/08/2017
---


# Range.Shading Property (Word)

Returns a  **Shading** object that refers to the shading formatting for the specified object.


## Syntax

 _expression_ . **Shading**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example applies yellow shading to the first paragraph in the selection.


```vb
With Selection.Paragraphs(1).Shading 
 .Texture = wdTexture12Pt5Percent 
 .BackgroundPatternColorIndex = wdYellow 
 .ForegroundPatternColorIndex = wdBlack 
End With
```

This example applies 10 percent shading to the first word in the active document.




```vb
ActiveDocument.Words(1).Shading.Texture = wdTexture10Percent
```


## See also


#### Concepts


[Range Object](range-object-word.md)

