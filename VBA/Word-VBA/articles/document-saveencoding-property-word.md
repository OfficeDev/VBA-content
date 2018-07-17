---
title: Document.SaveEncoding Property (Word)
keywords: vbawd10.chm158007629
f1_keywords:
- vbawd10.chm158007629
ms.prod: word
api_name:
- Word.Document.SaveEncoding
ms.assetid: 9a69851e-af52-d257-d887-c90bd98eeac0
ms.date: 06/08/2017
---


# Document.SaveEncoding Property (Word)

Returns or sets the encoding to use when saving a document. Read/write  **MsoEncoding** .


## Syntax

 _expression_ . **SaveEncoding**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example specifies Western encoding for saving the current document.


```vb
ActiveDocument.SaveEncoding = msoEncodingWestern
```


## See also


#### Concepts


[Document Object](document-object-word.md)

