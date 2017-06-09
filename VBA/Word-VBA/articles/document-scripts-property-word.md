---
title: Document.Scripts Property (Word)
keywords: vbawd10.chm158007616
f1_keywords:
- vbawd10.chm158007616
ms.prod: word
api_name:
- Word.Document.Scripts
ms.assetid: 5602a262-f4e2-bc9c-1457-68536adf7ac4
ms.date: 06/08/2017
---


# Document.Scripts Property (Word)

Returns a  **Scripts** collection that represents the collection of HTML scripts in the specified object.


## Syntax

 _expression_ . **Scripts**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example displays the text in the first  **Script** object of the active document.


```vb
Debug.Print ActiveDocument.Scripts(1).ScriptText
```


## See also


#### Concepts


[Document Object](document-object-word.md)

