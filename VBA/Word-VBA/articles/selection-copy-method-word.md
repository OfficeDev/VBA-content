---
title: Selection.Copy Method (Word)
keywords: vbawd10.chm158662776
f1_keywords:
- vbawd10.chm158662776
ms.prod: word
api_name:
- Word.Selection.Copy
ms.assetid: 5af32d69-5c0f-428a-44f3-35c75b5fb050
ms.date: 06/08/2017
---


# Selection.Copy Method (Word)

Copies the specified selection to the Clipboard.


## Syntax

 _expression_ . **Copy**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example copies the contents of the selection into a new document.


```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Copy 
 Documents.Add.Content.Paste 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

