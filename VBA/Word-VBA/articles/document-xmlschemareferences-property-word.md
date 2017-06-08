---
title: Document.XMLSchemaReferences Property (Word)
keywords: vbawd10.chm158007757
f1_keywords:
- vbawd10.chm158007757
ms.prod: word
api_name:
- Word.Document.XMLSchemaReferences
ms.assetid: 7008fb35-017d-2f14-0627-9b524138137c
ms.date: 06/08/2017
---


# Document.XMLSchemaReferences Property (Word)

Returns an XMLSchemaReferences collection that represents the schemas attached to a document.


## Syntax

 _expression_ . **XMLSchemaReferences**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Example

The following example reloads the first schema attached to the active document.


```vb
Dim objSchema As XMLSchemaReference 
 
Set objSchema = ActiveDocument.XMLSchemaReferences.Item(1) 
 
objSchema.Reload
```


## See also


#### Concepts


[Document Object](document-object-word.md)

