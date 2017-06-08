---
title: Conflicts.RejectAll Method (Word)
keywords: vbawd10.chm174391398
f1_keywords:
- vbawd10.chm174391398
ms.prod: word
api_name:
- Word.Conflicts.RejectAll
ms.assetid: bd3779d6-8cba-9cf8-d8ec-a9952e3918ad
ms.date: 06/08/2017
---


# Conflicts.RejectAll Method (Word)

Rejects all of the user's changes and retains the server copy of the document.


## Syntax

 _expression_ . **RejectAll**

 _expression_ An expression that returns a **Conflicts** object.


### Return Value

Nothing


## Example

The following code example rejects all the user's changes and retains the server copy of the active document.


```vb
ActiveDocument.CoAuthoring.Conflicts.RejectAll
```


## See also


#### Concepts


[Conflicts Object](conflicts-object-word.md)

