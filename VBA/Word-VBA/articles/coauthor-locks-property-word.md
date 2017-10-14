---
title: CoAuthor.Locks Property (Word)
keywords: vbawd10.chm81068036
f1_keywords:
- vbawd10.chm81068036
ms.prod: word
api_name:
- Word.CoAuthor.Locks
ms.assetid: 9f502e4e-2414-0232-78d0-5ce64d4297f0
ms.date: 06/08/2017
---


# CoAuthor.Locks Property (Word)

Returns a [CoAuthLocks](coauthlocks-object-word.md) collection that represents the locks in the document that are associated with the specified co author. Read-only.


## Syntax

 _expression_ . **Locks**

 _expression_ An expression that returns a **CoAuthor** object.


## Example

The following code example displays the number of locks that are associated with the first co author in the active document.


```vb
Dim lockCount As Integer 
Dim coAuth As CoAuthor 
 
Set coAuth = ActiveDocument.CoAuthoring.Authors(1) 
lockCount = coAuth.Locks.Count 
 
MsgBox "There are " &; lockCount &; _ 
 " locks in the active document for " &; _ 
 coAuth.Name &; "."
```


## See also


#### Concepts


[CoAuthor Object](coauthor-object-word.md)

