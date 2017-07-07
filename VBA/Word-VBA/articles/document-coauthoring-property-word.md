---
title: Document.CoAuthoring Property (Word)
keywords: vbawd10.chm158007896
f1_keywords:
- vbawd10.chm158007896
ms.prod: word
api_name:
- Word.Document.CoAuthoring
ms.assetid: b67ac270-c583-f141-bf86-6fc385987636
ms.date: 06/08/2017
---


# Document.CoAuthoring Property (Word)

Returns a [CoAuthoring](coauthoring-object-word.md) object that provides the entry point into the co authoring object model. Read-only.


## Syntax

 _expression_ . **CoAuthoring**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Remarks

The [CoAuthoring](coauthoring-object-word.md) object provides information about co authoring at the document level. For example, the[CoAuthoring](coauthoring-object-word.md) object can provide information about whether there are any locks in the document, which users have current locks on the document, or whether or not updates to the document content is available from the server. Use the **CoAuthoring** property to return the[CoAuthoring](coauthoring-object-word.md) object.


## Example

The following code example gets a reference to the [CoAuthoring](coauthoring-object-word.md) object through the **CoAuthoring** property of the active document.


```vb
Dim coAuth As CoAuthoring 
Set coAuth = ActiveDocument.CoAuthoring
```


## See also


#### Concepts


[Document Object](document-object-word.md)

