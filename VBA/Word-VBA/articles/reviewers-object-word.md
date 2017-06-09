---
title: Reviewers Object (Word)
keywords: vbawd10.chm3226
f1_keywords:
- vbawd10.chm3226
ms.prod: word
api_name:
- Word.Reviewers
ms.assetid: ae1bec96-e6dc-39f0-421a-dfeeb95c9049
ms.date: 06/08/2017
---


# Reviewers Object (Word)

A collection of  **[Reviewer](reviewer-object-word.md)** objects that represents the reviewers of one or more documents. The **Reviewers** collection contains the names of all reviewers who have reviewed documents opened or edited on a computer.


## Remarks

Use  **Reviewers** (Index), where Index is the name or index number of the reviewer, to return a single reviewer in the **Reviewers** collection. This example hides revisions made by the first reviewer in the **Reviewers** collection.


```vb
Sub HideAuthorRevisions(blnRev As Boolean) 
 ActiveWindow.View.Reviewers(Index:=1) _ 
 .Visible = False 
End Sub
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


