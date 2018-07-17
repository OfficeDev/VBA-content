---
title: CoAuthoring Object (Word)
keywords: vbawd10.chm3889
f1_keywords:
- vbawd10.chm3889
ms.prod: word
api_name:
- Word.CoAuthoring
ms.assetid: d36ac5a7-6479-6565-dbb0-969d06b31f30
ms.date: 06/08/2017
---


# CoAuthoring Object (Word)

Provides the primary entry point to the co authoring object model.


## Remarks

The  **CoAuthoring** object provides information about co authoring at the document level. For example, the **CoAuthoring** object can provide information about whether there are any locks in the document, which users have current locks in the document, or whether or not updates to the document content is available from the server. Use the **[CoAuthoring](document-coauthoring-property-word.md)** property to return the **CoAuthoring** object.


## Example

The following code example gets the number of locks in the active document.


```vb
Sub CountLocks() 
Dim i As Integer 
 
i = ActiveDocument.CoAuthoring.Locks.Count 
 
MsgBox i 
 
End Sub
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

