---
title: Selection.Cut Method (Word)
keywords: vbawd10.chm158662775
f1_keywords:
- vbawd10.chm158662775
ms.prod: word
api_name:
- Word.Selection.Cut
ms.assetid: 1e5dec1a-c621-2b54-ab7f-78ce90c0936f
ms.date: 06/08/2017
---


# Selection.Cut Method (Word)

Removes the specified object from the document and moves it to the Clipboard.


## Syntax

 _expression_ . **Cut**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

The contents of the selection are moved to the Clipboard but a collapsed selection remains in the document.


## Example

This example deletes the contents of the selection and pastes them into a new document.


```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Cut 
 Documents.Add.Content.Paste 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

