---
title: Document.Container Property (Word)
keywords: vbawd10.chm158007378
f1_keywords:
- vbawd10.chm158007378
ms.prod: word
api_name:
- Word.Document.Container
ms.assetid: f2a0ebbe-98dc-dfc4-5879-da2b79e75b7d
ms.date: 06/08/2017
---


# Document.Container Property (Word)

Returns the object that represents the container application for the specified document. Read-only  **Object** .


## Syntax

 _expression_ . **Container**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The  **Container** property provides access to the specified document's container application if the document is embedded in another application as an OLE object. This property also provides a pathway into the object model of the container application if a Word document is opened as an ActiveX document — for example, when a Word document is opened in Microsoft Office Binder or Internet Explorer.


## Example

This example displays the name of the container application for the first shape in the active document. For the example to work, this shape must be an OLE object.


```
Msgbox ActiveDocument.Shapes(1).OLEFormat.Object.Container.Name
```


## See also


#### Concepts


[Document Object](document-object-word.md)

