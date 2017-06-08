---
title: Comment.IsInk Property (Word)
keywords: vbawd10.chm154993652
f1_keywords:
- vbawd10.chm154993652
ms.prod: word
api_name:
- Word.Comment.IsInk
ms.assetid: 57204e17-cf5a-d006-0738-b1f1ef62632f
ms.date: 06/08/2017
---


# Comment.IsInk Property (Word)

Returns a  **Boolean** that represents whether a comment is a handwritten comment.


## Syntax

 _expression_ . **IsInk**

 _expression_ An expression that returns a **[Comment](comment-object-word.md)** object.


## Example

The following example removes all handwritten comments from the active document.


```vb
Dim objComment As Comment 
 
For Each objComment In ActiveDocument.Comments 
 If objComment.IsInk = True Then 
 objComment.Delete 
 End If 
Next
```


## See also


#### Concepts


[Comment Object](comment-object-word.md)

