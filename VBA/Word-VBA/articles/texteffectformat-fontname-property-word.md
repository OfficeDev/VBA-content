---
title: TextEffectFormat.FontName Property (Word)
keywords: vbawd10.chm164560999
f1_keywords:
- vbawd10.chm164560999
ms.prod: word
api_name:
- Word.TextEffectFormat.FontName
ms.assetid: fe1f6714-ed34-0c7f-c156-b91b601149de
ms.date: 06/08/2017
---


# TextEffectFormat.FontName Property (Word)

Returns or sets the name of the font for the dropped capital letter. Read/write  **String** .


## Syntax

 _expression_ . **FontName**

 _expression_ A variable that represents a **[TextEffectFormat](texteffectformat-object-word.md)** object.


## Example

This example sets Arial as the font for the dropped capital letter for the first paragraph in the active document.


```vb
With ActiveDocument.Paragraphs(1).DropCap 
 .FontName = "Arial" 
 .Position = wdDropNormal 
 .LinesToDrop = 3 
 .DistanceFromText = InchesToPoints(0.1) 
End With
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-word.md)

