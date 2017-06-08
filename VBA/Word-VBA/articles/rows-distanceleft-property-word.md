---
title: Rows.DistanceLeft Property (Word)
keywords: vbawd10.chm155975700
f1_keywords:
- vbawd10.chm155975700
ms.prod: word
api_name:
- Word.Rows.DistanceLeft
ms.assetid: 2798e176-d175-6504-7eac-4248bdb0ceb3
ms.date: 06/08/2017
---


# Rows.DistanceLeft Property (Word)

Returns or sets the distance (in points) between the document text and the left edge of the specified table. Read/write  **Single** .


## Syntax

 _expression_ . **DistanceLeft**

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

