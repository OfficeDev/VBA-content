---
title: Document.GridOriginFromMargin Property (Word)
keywords: vbawd10.chm158007604
f1_keywords:
- vbawd10.chm158007604
ms.prod: word
api_name:
- Word.Document.GridOriginFromMargin
ms.assetid: 137b250a-31d6-89c7-365b-285f14ae3dac
ms.date: 06/08/2017
---


# Document.GridOriginFromMargin Property (Word)

 **True** if Microsoft Word starts the character grid from the upper-left corner of the page. Read/write **Boolean** .


## Syntax

 _expression_ . **GridOriginFromMargin**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets Microsoft Word to start the character grid for the active document from the upper-left corner of the page.


```vb
ActiveDocument.GridOriginFromMargin = True
```


## See also


#### Concepts


[Document Object](document-object-word.md)

