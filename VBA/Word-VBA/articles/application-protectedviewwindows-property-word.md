---
title: Application.ProtectedViewWindows Property (Word)
keywords: vbawd10.chm158335466
f1_keywords:
- vbawd10.chm158335466
ms.prod: word
api_name:
- Word.Application.ProtectedViewWindows
ms.assetid: eb1c8cae-c0da-0a84-316e-808302869b26
ms.date: 06/08/2017
---


# Application.ProtectedViewWindows Property (Word)

Returns a [ProtectedViewWindows](protectedviewwindows-object-word.md) collection that represents all protected view windows. Read-only.


## Syntax

 _expression_ . **ProtectedViewWindows**

 _expression_ An expression that returns an **Application** object.


## Example

The following code example displays the number of protected view windows that are open.


```vb
MsgBox "There are " &; ProtectedViewWindows.Count &; _ 
 " protected view windows open."
```


## See also


#### Concepts


[Application Object](application-object-word.md)

