---
title: Document.SelectLinkedControls Method (Word)
keywords: vbawd10.chm158007845
f1_keywords:
- vbawd10.chm158007845
ms.prod: word
api_name:
- Word.Document.SelectLinkedControls
ms.assetid: cae4e00c-a34f-8581-07f9-b58722ec399e
ms.date: 06/08/2017
---


# Document.SelectLinkedControls Method (Word)

Returns a  **[ContentControls](contentcontrols-object-word.md)** collection that represents all content controls in a document that are linked to the specific custom XML node in the document's XML data store as specified by the Node parameter. Read-only.


## Syntax

 _expression_ . **SelectLinkedControls**( **_Node_** )

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Node_|Required| **CustomXMLNode**|The XML node in the document's data store to which the content controls are linked.|

### Return Value

ContentControls


## See also


#### Concepts


[Document Object](document-object-word.md)

