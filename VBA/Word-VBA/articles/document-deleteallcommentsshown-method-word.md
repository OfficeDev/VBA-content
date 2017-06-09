---
title: Document.DeleteAllCommentsShown Method (Word)
keywords: vbawd10.chm158007670
f1_keywords:
- vbawd10.chm158007670
ms.prod: word
api_name:
- Word.Document.DeleteAllCommentsShown
ms.assetid: b0cdbc8e-973c-1921-a646-d2f5ef091ce9
ms.date: 06/08/2017
---


# Document.DeleteAllCommentsShown Method (Word)

Deletes all revisions in a specified document that are displayed on the screen.


## Syntax

 _expression_ . **DeleteAllCommentsShown**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example hides all comments made by "Jeff Smith" and deletes all other displayed comments.


```vb
Sub HideDeleteComments() 
 Dim rev As Reviewer 
 With ActiveWindow.View 
 'Display all comments and revisions 
 .ShowRevisionsAndComments = True 
 .ShowFormatChanges = True 
 .ShowInsertionsAndDeletions = True 
 
 For Each rev In .Reviewers 
 rev.Visible = True 
 Next 
 
 'Hide only the revisions/comments made by the 
 'reviewer named "Jeff Smith" 
 .Reviewers(Index:="Jeff Smith").Visible = False 
 End With 
 
 'Delete all comments displayed in the active view 
 ActiveDocument.DeleteAllCommentsShown 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

