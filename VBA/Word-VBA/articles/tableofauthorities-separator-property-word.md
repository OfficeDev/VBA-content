---
title: TableOfAuthorities.Separator Property (Word)
keywords: vbawd10.chm152109061
f1_keywords:
- vbawd10.chm152109061
ms.prod: word
api_name:
- Word.TableOfAuthorities.Separator
ms.assetid: 4da467e9-77df-c656-ed37-f3388ba92b7c
ms.date: 06/08/2017
---


# TableOfAuthorities.Separator Property (Word)

Returns or sets up to five characters that appear between the sequence number and the page number in a table of authorities. Read/write  **String** .


## Syntax

 _expression_ . **Separator**

 _expression_ Required. A variable that represents a **[TableOfAuthorities](tableofauthorities-object-word.md)** collection.


## Remarks

This property corresponds to the \d switch for a Table of Authorities (TOA) field. A hyphen (-) is the default character.


## Example

This example inserts a table of authorities at the beginning of the active document, and then it formats the table to include a sequence number and a page number, separated by a hyphen (-).


```vb
Set myRange = ActiveDocument.Range(0, 0) 
With ActiveDocument.TablesOfAuthorities.Add(Range:=myRange) 
 .IncludeSequenceName = "Chapter" 
 .Separator = "-" 
End With
```


## See also


#### Concepts


[TableOfAuthorities Object](tableofauthorities-object-word.md)

