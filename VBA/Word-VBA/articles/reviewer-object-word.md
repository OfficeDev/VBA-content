---
title: Reviewer Object (Word)
keywords: vbawd10.chm1198
f1_keywords:
- vbawd10.chm1198
ms.prod: word
api_name:
- Word.Reviewer
ms.assetid: d7824ac4-d62a-b8f8-a80c-6999a999456c
ms.date: 06/08/2017
---


# Reviewer Object (Word)

Represents a single reviewer of a document in which changes have been tracked. The  **Reviewer** object is a member of the **[Reviewers](reviewers-object-word.md)** collection.


## Remarks

Use  **Reviewers** ( _Index_ ), where _Index_ is the name or number of the reviewer, to return a **Reviewer** object. Use the **Visible** property to display or hide individual reviewers in a document. The following code example hides the reviewer named "Jeff Smith" and displays the reviewer named "Judy Lew." This assumes that "Jeff Smith" and "Judy Lew" are members of the **Reviewers** collection. If they are not, you will receive an error.


```vb
Sub ShowHide() 
 With ActiveWindow.View 
 .Reviewers("Jeff Smith").Visible = False 
 .Reviewers("Judy Lew").Visible = True 
 End With 
End Sub
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


