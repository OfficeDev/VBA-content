---
title: DropCap.FontName Property (Word)
keywords: vbawd10.chm156631051
f1_keywords:
- vbawd10.chm156631051
ms.prod: word
api_name:
- Word.DropCap.FontName
ms.assetid: 5c89102e-fbf2-cb40-d89b-fbeb56386da1
ms.date: 06/08/2017
---


# DropCap.FontName Property (Word)

Returns or sets a  **String** that represents the name of the font for the dropped capital letter. Read/write.


## Syntax

 _expression_ . **FontName**

 _expression_ A variable that represents a **[DropCap](dropcap-object-word.md)** object.


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


[DropCap Object](dropcap-object-word.md)

