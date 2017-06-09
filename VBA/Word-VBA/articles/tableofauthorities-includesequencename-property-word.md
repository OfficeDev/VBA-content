---
title: TableOfAuthorities.IncludeSequenceName Property (Word)
keywords: vbawd10.chm152109062
f1_keywords:
- vbawd10.chm152109062
ms.prod: word
api_name:
- Word.TableOfAuthorities.IncludeSequenceName
ms.assetid: 15f3801c-4d79-c01f-4a67-5b09e1f14577
ms.date: 06/08/2017
---


# TableOfAuthorities.IncludeSequenceName Property (Word)

Returns or sets the Sequence (SEQ) field identifier for a table of authorities. Read/write  **String** .


## Syntax

 _expression_ . **IncludeSequenceName**

 _expression_ An expression that returns a **[TableOfAuthorities](tableofauthorities-object-word.md)** object.


## Remarks

Corresponds to the \s switch for a Table of Authorities (TOA) field.


## Example

This example inserts a table of authorities at the beginning of the active document and then formats the table to include the Chapter sequence field number before the page number (for example, "Chapter 2-14").


```vb
Dim rngTemp As Range 
Dim toaTemp As TableOfAuthorities 
 
Set rngTemp = ActiveDocument.Range(Start:=0, End:=0) 
Set toaTemp = _ 
 ActiveDocument.TablesOfAuthorities.Add(Range:=rngTemp) 
 
toaTemp.IncludeSequenceName = "Chapter"
```

This example returns the sequence name for the first table of authorities.




```vb
Dim strSequence As String 
 
strSequence = _ 
 ActiveDocument.TablesOfAuthorities(1).IncludeSequenceName
```


## See also


#### Concepts


[TableOfAuthorities Object](tableofauthorities-object-word.md)

