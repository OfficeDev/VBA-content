---
title: TableOfAuthorities.PageRangeSeparator Property (Word)
keywords: vbawd10.chm152109064
f1_keywords:
- vbawd10.chm152109064
ms.prod: word
api_name:
- Word.TableOfAuthorities.PageRangeSeparator
ms.assetid: f2b2c68f-15b2-5eb1-1af2-981920f18cc7
ms.date: 06/08/2017
---


# TableOfAuthorities.PageRangeSeparator Property (Word)

Returns or sets the characters (up to five) that separate a range of pages in a table of authorities. Read/write  **String** .


## Syntax

 _expression_ . **PageRangeSeparator**

 _expression_ An expression that returns a **[TableOfAuthorities](tableofauthorities-object-word.md)** object.


## Remarks

The default is an en dash. Corresponds to the \g switch for a Table of Authorities (TOA) field. 


## Example

This example formats the first table of authorities in the active document to use a hyphen with a space on either side as the page range separator (for example, "9 - 12").


```vb
ActiveDocument.TablesOfAuthorities(1).PageRangeSeparator = " - "
```


## See also


#### Concepts


[TableOfAuthorities Object](tableofauthorities-object-word.md)

