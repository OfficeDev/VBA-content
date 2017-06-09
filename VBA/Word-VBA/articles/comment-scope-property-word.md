---
title: Comment.Scope Property (Word)
keywords: vbawd10.chm154993645
f1_keywords:
- vbawd10.chm154993645
ms.prod: word
api_name:
- Word.Comment.Scope
ms.assetid: 07ef4a30-9a3a-aed1-5c38-7f091ea3150b
ms.date: 06/08/2017
---


# Comment.Scope Property (Word)

Returns a  **[Range](range-object-word.md)** object that represents the range of text marked by the specified comment.


## Syntax

 _expression_ . **Scope**

 _expression_ An expression that returns a **[Comment](comment-object-word.md)** object.


## Example

This example displays the text associated with the first comment in the selection.


```vb
If Selection.Comments.Count >= 1 Then 
 Set myRange = Selection.Comments(1).Scope 
 MsgBox myRange.Text 
End If
```

This example copies the text associated with the last comment in the active document.




```vb
total = ActiveDocument.Comments.Count 
If total >= 1 Then ActiveDocument.Comments(total).Scope.Copy
```


## See also


#### Concepts


[Comment Object](comment-object-word.md)

