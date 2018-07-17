---
title: CoAuthUpdate Object (Word)
ms.prod: word
api_name:
- Word.CoAuthUpdate
ms.assetid: c00e5029-2e4b-97c0-33d3-86fdc53df535
ms.date: 06/08/2017
---


# CoAuthUpdate Object (Word)

Represents a range of text that has been updated by a co author.


## Remarks

When a document that has co authoring enabled is edited by more than one author, changes to the document by one author are pushed to other authors' versions of the document by using updates. When a co author performs an explicit document save (by pressing  **CTRL** + **S**, for example), changes made by other co authors are merged into the document as updates. The  **CoAuthUpdates** collection contains all changes that were merged into the document, where each change is a single update represented by a **CoAuthUpdate** object.

The contents of the  **CoAuthUpdates** collection remains the same until a co author performs another explicit document save. When the co author saves the document again, if there are no new changes from other co authors that are merged into the document, the **CoAuthUpdates** collection retains the same updates that were merged at the previous explicit save. If there are new changes that are merged into the document, the **CoAuthUpdates** collection contains the new updates for the document. Use a **CoAuthUpdate** object to retrieve an individual update from the **[CoAuthUpdates](http://msdn.microsoft.com/library/4a164415-0c6c-213b-da94-744e2394d1ef%28Office.15%29.aspx)** collection.


## Example

The following code example gets the associated text in the range of each  **CoAuthUpdate** object in the active document.


```vb
Dim caUpdate As CoAuthUpdate 
Dim strText As String 
 
For Each caUpdate In ActiveDocument.CoAuthoring.Updates 
    strText = caUpdate.Range.Text 
    MsgBox strText 
Next caUpdate
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


