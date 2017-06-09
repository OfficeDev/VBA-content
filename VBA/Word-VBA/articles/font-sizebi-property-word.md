---
title: Font.SizeBi Property (Word)
keywords: vbawd10.chm156369058
f1_keywords:
- vbawd10.chm156369058
ms.prod: word
api_name:
- Word.Font.SizeBi
ms.assetid: 521dfc53-1076-ace0-c5d4-7218c985eb7c
ms.date: 06/08/2017
---


# Font.SizeBi Property (Word)

Returns or sets the font size in points. Read/write  **Single** .


## Syntax

 _expression_ . **SizeBi**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

The  **SizeBi** property applies to text in a right-to-left language.


## Example

This example sets the font size of the first word to 20 points.


```vb
With ActiveDocument.Paragraphs(1).Range 
 .Words(1).Font.SizeBi = 20 
End With
```


## See also


#### Concepts


[Font Object](font-object-word.md)

