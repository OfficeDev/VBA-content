---
title: Version Object (Word)
keywords: vbawd10.chm2484
f1_keywords:
- vbawd10.chm2484
ms.prod: word
api_name:
- Word.Version
ms.assetid: 63eeefb0-2d63-75e6-a070-a4a80f243bc4
ms.date: 06/08/2017
---


# Version Object (Word)

Represents a single version of a document. The  **Version** object is a member of the **Versions** collection. The **Versions** collection includes all the versions of the specified document.


## Remarks

Use  **Versions** (Index), where Index is the index number, to return a single **Version** object. The index number represents the position of the version in the **Versions** collection. The first version added to the **Versions** collection is index number 1. The following example displays the comment, author, and date of the first version of the active document.


```vb
If ActiveDocument.Versions.Count >= 1 Then 
 With ActiveDocument.Versions(1) 
 MsgBox "Comment = " &; .Comment &; vbCr &; "Author = " &; _ 
 .SavedBy &; vbCr &; "Date = " &; .Date 
 End With 
End If
```

Use the  **Save** method to add an item to the **Versions** collection. The following example adds a version of the active document with the specified comment.




```vb
ActiveDocument.Versions.Save _ 
 Comment:="incorporated Judy's revisions"
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


