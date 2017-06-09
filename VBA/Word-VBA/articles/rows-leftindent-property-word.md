---
title: Rows.LeftIndent Property (Word)
keywords: vbawd10.chm155975689
f1_keywords:
- vbawd10.chm155975689
ms.prod: word
api_name:
- Word.Rows.LeftIndent
ms.assetid: bb5ee915-a41a-e447-7326-b6b6e0e2d6d2
ms.date: 06/08/2017
---


# Rows.LeftIndent Property (Word)

Returns or sets a  **Single** that represents the left indent value (in points) for the specified table rows. Read/write.


## Syntax

 _expression_ . **LeftIndent**

 _expression_ A variable that represents a **[Rows](rows-object-word.md)** collection.


## Example

This example sets the left indent for all rows in the first table in the active document.


```vb
ActiveDocument.Tables(1).Rows.LeftIndent = InchesToPoints(1)
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

