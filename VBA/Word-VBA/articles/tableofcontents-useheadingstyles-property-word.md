---
title: TableOfContents.UseHeadingStyles Property (Word)
keywords: vbawd10.chm152240129
f1_keywords:
- vbawd10.chm152240129
ms.prod: word
api_name:
- Word.TableOfContents.UseHeadingStyles
ms.assetid: c026c00b-f3ec-b350-d046-0761b6e70851
ms.date: 06/08/2017
---


# TableOfContents.UseHeadingStyles Property (Word)

 **True** if built-in heading styles are used to create a table of contents. Read/write **Boolean** .


## Syntax

 _expression_ . **UseHeadingStyles**

 _expression_ Required. A variable that represents a **[TableOfContents](tableofcontents-object-word.md)** collection.


## Example

This example formats the first table of contents in the active document to compile entries formatted with the Heading 1, Heading 2, or Heading 3 style.


```vb
If ActiveDocument.TablesOfContents.Count >= 1 Then 
 With ActiveDocument.TablesOfContents(1) 
 .UseHeadingStyles = True 
 .UseFields = False 
 .UpperHeadingLevel = 1 
 .LowerHeadingLevel = 3 
 End With 
End If
```


## See also


#### Concepts


[TableOfContents Object](tableofcontents-object-word.md)

