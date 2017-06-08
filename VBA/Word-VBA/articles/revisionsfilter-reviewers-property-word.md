---
title: RevisionsFilter.Reviewers Property (Word)
keywords: vbawd10.chm40566787
f1_keywords:
- vbawd10.chm40566787
ms.prod: word
ms.assetid: c076a572-602e-537a-52ce-eb36b778ad73
ms.date: 06/08/2017
---


# RevisionsFilter.Reviewers Property (Word)

Returns a [Reviewers](reviewers-object-word.md) object that represents the collection of reviewers of one or more documents.


## Syntax

 _expression_ . **Reviewers**

 _expression_ A variable that represents a **RevisionsFilter** object.


## Remarks

The  **Reviewers** collection returned by **Reviewers** contains the names of all reviewers who have reviewed documents opened or edited on a computer.


 **Note**  The  **Reviewers** property replaces the **View.Reviewers** property found in previous versions of Word, which is now deprecated.


## Example

This example shows how to get the count of all reviewers in the document in the active window. This example assumes that the document in the active window contains revisions made by one or more reviewers.


```vb
Public Sub Reviewers_Example()

   Debug.Print ActiveWindow.View.RevisionsFilter.Reviewers.Count

End Sub
```


## Property value

 **REVIEWERS**


## See also


#### Other resources


[RevisionsFilter Object](revisionsfilter-object-word.md)


