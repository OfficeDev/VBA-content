---
title: TableOfContents.HidePageNumbersInWeb Property (Word)
keywords: vbawd10.chm152240140
f1_keywords:
- vbawd10.chm152240140
ms.prod: word
api_name:
- Word.TableOfContents.HidePageNumbersInWeb
ms.assetid: 81d77980-099e-e048-b219-d10b64cd6a38
ms.date: 06/08/2017
---


# TableOfContents.HidePageNumbersInWeb Property (Word)

Returns or sets whether page numbers in a table of contents or a table of figures should be hidden when publishing to the Web. Read/write  **Boolean** .


## Syntax

 _expression_ . **HidePageNumbersInWeb**

 _expression_ A variable that represents a **[TableOfContents](tableofcontents-object-word.md)** collection.


## Example

This example hides page numbers in the first table of contents if the document is to be published to the Web.


```vb
ActiveDocument.TableOfContents(1).HidePageNumbersInWeb = True
```


## See also


#### Concepts


[TableOfContents Object](tableofcontents-object-word.md)

