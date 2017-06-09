---
title: TableOfContents.UseHyperlinks Property (Word)
keywords: vbawd10.chm152240139
f1_keywords:
- vbawd10.chm152240139
ms.prod: word
api_name:
- Word.TableOfContents.UseHyperlinks
ms.assetid: 2ff74d58-6411-eb10-1ce4-86d0b8e37490
ms.date: 06/08/2017
---


# TableOfContents.UseHyperlinks Property (Word)

Returns or sets whether entries in a table of contents should be formatted as hyperlinks when publishing to the Web. Read/write  **Boolean** .


## Syntax

 _expression_ . **UseHyperlinks**

 _expression_ Required. A variable that represents a **[TableOfContents](tableofcontents-object-word.md)** collection.


## Example

This example formats the first table of contents in the document using hyperlinks.


```vb
ActiveDocument.TableOfContents(1).UseHyperlinks = True
```


## See also


#### Concepts


[TableOfContents Object](tableofcontents-object-word.md)

