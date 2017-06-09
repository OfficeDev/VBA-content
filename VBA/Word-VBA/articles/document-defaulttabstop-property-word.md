---
title: Document.DefaultTabStop Property (Word)
keywords: vbawd10.chm158007344
f1_keywords:
- vbawd10.chm158007344
ms.prod: word
api_name:
- Word.Document.DefaultTabStop
ms.assetid: 55c7a9e4-0a25-cd32-36b0-fc9431b1d110
ms.date: 06/08/2017
---


# Document.DefaultTabStop Property (Word)

Returns or sets the interval (in points) between the default tab stops in the specified document. Read/write  **Single** .


## Syntax

 _expression_ . **DefaultTabStop**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets the default tab stops in the active document to 1 inch. The  **[InchesToPoints](application-inchestopoints-method-word.md)** method is used to convert inches to points.


```vb
ActiveDocument.DefaultTabStop = InchesToPoints(1)
```


## See also


#### Concepts


[Document Object](document-object-word.md)

