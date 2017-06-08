---
title: CoAuthoring.PendingUpdates Property (Word)
keywords: vbawd10.chm254869507
f1_keywords:
- vbawd10.chm254869507
ms.prod: word
api_name:
- Word.CoAuthoring.PendingUpdates
ms.assetid: ddc669ca-89dd-d321-4544-cc24e18270c6
ms.date: 06/08/2017
---


# CoAuthoring.PendingUpdates Property (Word)

Returns  **true** if the document has pending updates that have not been accepted. Read-only.


## Syntax

 _expression_ . **PendingUpdates**

 _expression_ An expression that returns a **[CoAuthoring](coauthoring-object-word.md)** object.


## Example

The following code example displays a message that indicates whether content updates are pending for the active document.


```vb
If ActiveDocument.CoAuthoring.PendingUpdates Then 
MsgBox "There are content updates pending." 
Else: MsgBox "There are no pending updates." 
End If
```


## See also


#### Concepts


[CoAuthoring Object](coauthoring-object-word.md)

