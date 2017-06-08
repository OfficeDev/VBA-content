---
title: Rows.DistanceBottom Property (Word)
keywords: vbawd10.chm155975694
f1_keywords:
- vbawd10.chm155975694
ms.prod: word
api_name:
- Word.Rows.DistanceBottom
ms.assetid: 21d0bb53-69d5-d579-a7eb-690e8f2742fb
ms.date: 06/08/2017
---


# Rows.DistanceBottom Property (Word)

Returns or sets the distance (in points) between the document text and the bottom edge of the specified table. Read/write  **Single** .


## Syntax

 _expression_ . **DistanceBottom**

 _expression_ A variable that represents a **[Rows](rows-object-word.md)** collection.


## Remarks

This property doesn't have any effect if  **WrapAroundText** is **False** .


## Example

This example sets text to wrap around the first table in the active document and sets the distance for wrapped text to 20 points on all sides of the table.


```vb
With ActiveDocument.Tables(1).Rows 
 .WrapAroundText = True 
 .DistanceLeft = 20 
 .DistanceRight = 20 
 .DistanceTop = 20 
 .DistanceBottom = 20 
End With
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

