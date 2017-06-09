---
title: Selection.ShrinkDiscontiguousSelection Method (Word)
keywords: vbawd10.chm158663675
f1_keywords:
- vbawd10.chm158663675
ms.prod: word
api_name:
- Word.Selection.ShrinkDiscontiguousSelection
ms.assetid: ce703cb4-8a20-b59d-ccf7-c0c93327a9ca
ms.date: 06/08/2017
---


# Selection.ShrinkDiscontiguousSelection Method (Word)

Cancels the selection of all but the most recently selected text when a selection contains multiple, unconnected selections.


## Syntax

 _expression_ . **ShrinkDiscontiguousSelection**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example cancels the selection of all but the most recently selected text and formats with bold and small caps the text remaining in the selection. This example assumes there are multiple selections in the document.


```vb
Sub ShrinkMultipleSelection() 
 With Selection 
 .ShrinkDiscontiguousSelection 
 .Font.Bold = True 
 .Font.SmallCaps = True 
 End With 
End Sub
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

