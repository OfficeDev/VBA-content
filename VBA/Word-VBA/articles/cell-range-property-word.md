---
title: Cell.Range Property (Word)
keywords: vbawd10.chm156106752
f1_keywords:
- vbawd10.chm156106752
ms.prod: word
api_name:
- Word.Cell.Range
ms.assetid: 579a25ad-91fa-a7c9-7eb8-4307521aeddd
ms.date: 06/08/2017
---


# Cell.Range Property (Word)

Returns a  **[Range](range-object-word.md)** object that represents the portion of a document that's contained in the specified object.


## Syntax

 _expression_ . **Range**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


## Example

This example copies thecontents of the first cell in the first row in the first table.


```vb
If ActiveDocument.Tables.Count >= 1 Then _ 
 ActiveDocument.Tables(1).Rows(1).Cells(1).Range.Copy
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

