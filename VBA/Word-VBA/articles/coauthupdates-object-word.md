---
title: CoAuthUpdates Object (Word)
ms.assetid: afd0abeb-276e-96f4-ee8a-01f263e69121
ms.prod: word
ms.date: 06/08/2017
---


# CoAuthUpdates Object (Word)

A collection of [CoAuthUpdate](coauthupdate-object-word.md) objects that represent the updates that were merged into the document at the last explicit save.


## Remarks

When a document with co authoring enabled is edited by more than one author, changes to the document by one author are pushed to other authors' versions of the document using updates. When a co author performs an explicit document save (by pressing  **CTRL** + **S**, for example), changes made by other co authors are merged into the document as updates. The  **CoAuthUpdates** collection contains all the changes that were merged into the document, where each change is a single update.

The contents of the  **CoAuthUpdates** collection remains the same until a co author performs another explicit document save. When the co author saves the document again, if there are no new changes from other co authors that are merged into the document, the **CoAuthUpdates** collection retains the same updates that were merged at the previous explicit save. If there are new changes that are merged into the document, the **CoAuthUpdates** collection contains the new updates for the document.


## Example

The following code example gets the number of the latest updates that were merged into the document at the last explicit save.


```vb
Dim countOfUpdates As Integer 
 
countOfUpdates = ActiveDocument.CoAuthoring.Updates.Count 
 
MsgBox "The number of updates is " &; countOfUpdates
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

