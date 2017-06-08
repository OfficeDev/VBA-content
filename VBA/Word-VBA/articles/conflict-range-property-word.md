---
title: Conflict.Range Property (Word)
keywords: vbawd10.chm78708739
f1_keywords:
- vbawd10.chm78708739
ms.prod: word
api_name:
- Word.Conflict.Range
ms.assetid: 8f3eb9c1-041e-62e0-d3f8-b9983f94ed9c
ms.date: 06/08/2017
---


# Conflict.Range Property (Word)

 Returns a[Range](range-object-word.md) object that represents the portion of a document that is contained in the specified object. Read-only.


## Syntax

 _expression_ . **Range**

 _expression_ An expression that returns a **Conflict** object.


## Example

The following code example returns the range associated with the second conflict in the active document.


```vb
Dim rng As Range 
 
Set rng = ActiveDocument.CoAuthoring.Conflicts(2).Range 

```


## See also


#### Concepts


[Conflict Object](conflict-object-word.md)

