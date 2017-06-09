---
title: Document.Permission Property (Word)
keywords: vbawd10.chm158007749
f1_keywords:
- vbawd10.chm158007749
ms.prod: word
api_name:
- Word.Document.Permission
ms.assetid: 17a100a0-3dc4-b15d-fcb6-e7bc57d08fc6
ms.date: 06/08/2017
---


# Document.Permission Property (Word)

Returns a  **Permission** object that represents the permission settings in the specified document.


## Syntax

 _expression_ . **Permission**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Example

The following example returns the permission settings for the active document.


```vb
Dim objPermission As Permission 
 
Set objPermission = ActiveDocument.Permission
```


## See also


#### Concepts


[Document Object](document-object-word.md)

