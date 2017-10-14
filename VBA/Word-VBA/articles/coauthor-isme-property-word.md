---
title: CoAuthor.IsMe Property (Word)
keywords: vbawd10.chm81068035
f1_keywords:
- vbawd10.chm81068035
ms.prod: word
api_name:
- Word.CoAuthor.IsMe
ms.assetid: bf6b8282-e114-8b6f-9e89-3bd93662d84e
ms.date: 06/08/2017
---


# CoAuthor.IsMe Property (Word)

Returns true if this author represents the current user. Read-only. 


## Syntax

 _expression_ . **IsMe**

 _expression_ An expression that returns a **CoAuthor** object.


## Example

The following code example checks the active document to see if the first co author in the CoAuthors collection is the current user.


```vb
If ActiveDocument.CoAuthoring.Authors(1).IsMe Then 
MsgBox "The current user is the first coauthor." 
End If
```


## See also


#### Concepts


[CoAuthor Object](coauthor-object-word.md)

