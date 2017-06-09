---
title: Document.SelectUnlinkedControls Method (Word)
keywords: vbawd10.chm158007846
f1_keywords:
- vbawd10.chm158007846
ms.prod: word
api_name:
- Word.Document.SelectUnlinkedControls
ms.assetid: 6d757837-0959-6754-bfae-e840ea7de339
ms.date: 06/08/2017
---


# Document.SelectUnlinkedControls Method (Word)

Returns a  **[ContentControls](contentcontrols-object-word.md)** collection that represents all content controls in a document that are not linked to an XML node in the document's XML data store. Read-only.


## Syntax

 _expression_ . **SelectUnlinkedControls**( **_Stream_** )

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Stream_|Optional| **CustomXMLPart**|A custom XML part reference. Setting this parameter filters the returned content controls to include only content controls that reference this  **CustomXMLPart** in their **XMLMapping** definition.|

### Return Value

ContentControls


## See also


#### Concepts


[Document Object](document-object-word.md)

