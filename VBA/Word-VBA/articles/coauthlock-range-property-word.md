---
title: CoAuthLock.Range Property (Word)
keywords: vbawd10.chm260046851
f1_keywords:
- vbawd10.chm260046851
ms.prod: word
api_name:
- Word.CoAuthLock.Range
ms.assetid: 092cafbc-09b1-75b7-660e-85b6cd2b5ba2
ms.date: 06/08/2017
---


# CoAuthLock.Range Property (Word)

Returns a [Range](range-object-word.md) object that represents the portion of a document that is contained in the specified object. Read-only.


## Syntax

 _expression_ . **Range**

 _expression_ An expression that returns a **[CoAuthLock](coauthlock-object-word.md)** object.


## Example

The following code example gets the document range for the first lock in the active document and displays the range text to the user.


```vb
MsgBox ActiveDocument.Coauthoring.Locks(1).Range
```


## See also


#### Concepts


[CoAuthLock Object](coauthlock-object-word.md)

