---
title: CoAuthors.Count Property (Word)
keywords: vbawd10.chm179961857
f1_keywords:
- vbawd10.chm179961857
ms.prod: word
api_name:
- Word.CoAuthors.Count
ms.assetid: 452917e0-133f-9bba-0e17-041370e0cb12
ms.date: 06/08/2017
---


# CoAuthors.Count Property (Word)

Returns the number of items in the [CoAuthors](coauthors-object-word.md) collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ An expression that returns a **CoAuthors** object.


## Example

The following code example displays the number of co authors in the active document.


```vb
MsgBox "The active document contains " &; _ 
 ActiveDocument.CoAuthoring.Authors.Count &; " authors."
```


## See also


#### Concepts


[CoAuthors Object](coauthors-object-word.md)

