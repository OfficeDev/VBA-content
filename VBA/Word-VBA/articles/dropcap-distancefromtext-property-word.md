---
title: DropCap.DistanceFromText Property (Word)
keywords: vbawd10.chm156631053
f1_keywords:
- vbawd10.chm156631053
ms.prod: word
api_name:
- Word.DropCap.DistanceFromText
ms.assetid: 8b86b00c-fc38-6fb3-8877-cba1eec814d7
ms.date: 06/08/2017
---


# DropCap.DistanceFromText Property (Word)

Returns or sets a  **Single** that represents the distance (in points) between the dropped capital letter and the paragraph text. Read/write.


## Syntax

 _expression_ . **DistanceFromText**

 _expression_ A variable that represents a **[DropCap](dropcap-object-word.md)** object.


## Example

This example sets a dropped capital letter for the first paragraph in the active document. The offset for the dropped capital letter is then set to 12 points.


```vb
With ActiveDocument.Paragraphs(1).DropCap 
 .Enable 
 .FontName= "Arial" 
 .Position = wdDropNormal 
 .DistanceFromText = 12 
End With
```


## See also


#### Concepts


[DropCap Object](dropcap-object-word.md)

