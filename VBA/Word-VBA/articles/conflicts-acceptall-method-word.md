---
title: Conflicts.AcceptAll Method (Word)
keywords: vbawd10.chm174391397
f1_keywords:
- vbawd10.chm174391397
ms.prod: word
api_name:
- Word.Conflicts.AcceptAll
ms.assetid: 8ccb2b0c-77ca-ff27-1e8d-5c1d504d8821
ms.date: 06/08/2017
---


# Conflicts.AcceptAll Method (Word)

Accepts all of the user's changes, removes the conflicts, and merges the changes into the server copy of the document.


## Syntax

 _expression_ . **AcceptAll**

 _expression_ An expression that returns a **Conflicts** object.


### Return Value

Nothing


## Example

The following code example accepts all the user changes for the conflicts in the active document.


```vb
ActiveDocument.CoAuthoring.Conflicts.AcceptAll
```


## See also


#### Concepts


[Conflicts Object](conflicts-object-word.md)

