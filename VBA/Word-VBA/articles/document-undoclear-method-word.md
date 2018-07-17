---
title: Document.UndoClear Method (Word)
keywords: vbawd10.chm158007550
f1_keywords:
- vbawd10.chm158007550
ms.prod: word
api_name:
- Word.Document.UndoClear
ms.assetid: 4ff5856a-ee8d-a9c8-a0a5-1d9c0a0dc9e9
ms.date: 06/08/2017
---


# Document.UndoClear Method (Word)

Clears the list of actions that can be undone for the specified document.


## Syntax

 _expression_ . **UndoClear**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

This method corresponds to the list of items that appears when you click the arrow beside the  **Undo** button on the **Standard** toolbar. Include this method at the end of a macro to keep Visual Basic actions from appearing in the **Undo** box (for example, "VBA-Selection.InsertAfter").


## Example

This example clears the list of actions that can be undone for the active document.


```vb
ActiveDocument.UndoClear
```


## See also


#### Concepts


[Document Object](document-object-word.md)

