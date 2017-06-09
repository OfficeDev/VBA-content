---
title: Document.GridSpaceBetweenVerticalLines Property (Word)
keywords: vbawd10.chm158007603
f1_keywords:
- vbawd10.chm158007603
ms.prod: word
api_name:
- Word.Document.GridSpaceBetweenVerticalLines
ms.assetid: 83658d56-6724-3e34-57bb-0b9cab537985
ms.date: 06/08/2017
---


# Document.GridSpaceBetweenVerticalLines Property (Word)

Returns or sets the interval at which Microsoft Word displays vertical character gridlines in print layout view. Read/write  **Long** .


## Syntax

 _expression_ . **GridSpaceBetweenVerticalLines**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets Microsoft Word to display every other vertical character gridline.


```vb
ActiveDocument.GridSpaceBetweenVerticalLines = 2
```


## See also


#### Concepts


[Document Object](document-object-word.md)

