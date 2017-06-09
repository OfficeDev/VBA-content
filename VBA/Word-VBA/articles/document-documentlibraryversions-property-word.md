---
title: Document.DocumentLibraryVersions Property (Word)
keywords: vbawd10.chm158007772
f1_keywords:
- vbawd10.chm158007772
ms.prod: word
api_name:
- Word.Document.DocumentLibraryVersions
ms.assetid: 1be5fae8-0ea1-115f-3786-6979a473448b
ms.date: 06/08/2017
---


# Document.DocumentLibraryVersions Property (Word)

Returns a  **DocumentLibraryVersions** collection that represents the collection of versions of a shared document that has versioning enabled and that is stored in a document library on a server.


## Syntax

 _expression_ . **DocumentLibraryVersions**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Example

The following example returns the collection of versions for the active document. This example assumes that the active document has versioning enabled and is stored in a shared document library on a server.


```vb
Dim objVersions As DocumentLibraryVersions 
 
Set objVersions = ActiveDocument.DocumentLibraryVersions
```


## See also


#### Concepts


[Document Object](document-object-word.md)

