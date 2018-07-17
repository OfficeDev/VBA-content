---
title: Document.AcceptAllRevisionsShown Method (Word)
keywords: vbawd10.chm158007668
f1_keywords:
- vbawd10.chm158007668
ms.prod: word
api_name:
- Word.Document.AcceptAllRevisionsShown
ms.assetid: bd9634cf-239a-2543-3681-579d4dd2f202
ms.date: 06/08/2017
---


# Document.AcceptAllRevisionsShown Method (Word)

Accepts all revisions in the specified document that are displayed on the screen.


## Syntax

 _expression_ . **AcceptAllRevisionsShown**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Use the  **RejectAllRevisionsShown** method to reject all revisions in a specified document that are displayed on the screen.


## Example

This example accepts all revisions displayed after hiding revisions made by "Jeff Smith." This example assumes that the active document was reviewed by more than one person and that the name of one of the reviewers is "Jeff Smith."


```vb
Sub AcceptAllChanges() 
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
 
 'Accept all revisions displayed in the active view 
 ActiveDocument.AcceptAllRevisionsShown 
 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

