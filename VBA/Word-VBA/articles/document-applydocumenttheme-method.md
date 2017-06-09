---
title: Document.ApplyDocumentTheme Method
keywords: vbawd10.chm158007842
f1_keywords:
- vbawd10.chm158007842
ms.prod: word
api_name:
- Word.ApplyDocumentTheme
ms.assetid: fd376134-f6d4-b6da-8eae-671e7e3b05e0
ms.date: 06/08/2017
---


# Document.ApplyDocumentTheme Method

Applies a document theme to a document.


## Syntax

 _expression_ . **ApplyDocumentTheme**( **_FileName_** )

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The path and file name of the theme to apply.|

## Example

The following example applies the Verve document theme to the active document.


```vb
ActiveDocument.ApplyDocumentTheme _ 
 "C:\Program Files\Microsoft Office\" &; _ 
 "Document Themes 12\Verve.thmx"
```


