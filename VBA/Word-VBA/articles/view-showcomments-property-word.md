---
title: View.ShowComments Property (Word)
keywords: vbawd10.chm161808419
f1_keywords:
- vbawd10.chm161808419
ms.prod: word
api_name:
- Word.View.ShowComments
ms.assetid: 01d688e8-9a5c-acd3-1626-d45a24a6b3b2
ms.date: 06/08/2017
---


# View.ShowComments Property (Word)

 **True** for Microsoft Word to display the comments in a document. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowComments**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Remarks

If revision marks are displayed in balloons in the right or left margin, comments are displayed in balloons. If revision marks are displayed inline, the text to which comments apply is surrounded by wide-angled square brackets; when a user places the mouse pointer over the text within the brackets, the related comment is displayed in a square balloon directly above the mouse pointer.


## Example

This example hides the comments in the active document. This example assumes that the document in the active window contains one or more comments.


```vb
Sub HideComments() 
 ActiveWindow.View.ShowComments = False 
End Sub
```


## See also


#### Concepts


[View Object](view-object-word.md)

