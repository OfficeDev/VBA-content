---
title: Document.TextLineEnding Property (Word)
keywords: vbawd10.chm158007654
f1_keywords:
- vbawd10.chm158007654
ms.prod: word
api_name:
- Word.Document.TextLineEnding
ms.assetid: 6e1f2243-473c-0294-623e-c09588645ee3
ms.date: 06/08/2017
---


# Document.TextLineEnding Property (Word)

Returns or sets a  **WdLineEndingType** constant indicating how Microsoft Word marks the line and paragraph breaks in documents saved as text files. Read/write.


## Syntax

 _expression_ . **TextLineEnding**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets the active document to enter a carriage return for line and paragraph breaks when it is saved as a text file.


```vb
Sub LineEndings() 
 ActiveDocument.TextLineEnding = wdCROnly 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

