---
title: TableOfAuthorities.PageNumberSeparator Property (Word)
keywords: vbawd10.chm152109066
f1_keywords:
- vbawd10.chm152109066
ms.prod: word
api_name:
- Word.TableOfAuthorities.PageNumberSeparator
ms.assetid: dad85c42-e819-1067-552e-e4387f344570
ms.date: 06/08/2017
---


# TableOfAuthorities.PageNumberSeparator Property (Word)

Returns of sets the characters (up to five) that separate individual page references in a table of authorities. Read/write  **String** .


## Syntax

 _expression_ . **PageNumberSeparator**

 _expression_ An expression that returns a **[TableOfAuthorities](tableofauthorities-object-word.md)** object.


## Remarks

The default is a comma and a space. Corresponds to the \l switch for a Table of Authorities (TOA) field.


## Example

This example formats the tables of authorities in the active document to use a comma as the page separator (for example, "9,12").


```vb
For Each myTOA In ActiveDocument.TablesOfAuthorities 
 myTOA.PageNumberSeparator = "," 
Next myTOA
```


## See also


#### Concepts


[TableOfAuthorities Object](tableofauthorities-object-word.md)

