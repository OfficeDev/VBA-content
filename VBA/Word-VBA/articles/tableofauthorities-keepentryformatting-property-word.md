---
title: TableOfAuthorities.KeepEntryFormatting Property (Word)
keywords: vbawd10.chm152109058
f1_keywords:
- vbawd10.chm152109058
ms.prod: word
api_name:
- Word.TableOfAuthorities.KeepEntryFormatting
ms.assetid: f8fcf3c1-0a72-071f-ee2c-341a78a43d36
ms.date: 06/08/2017
---


# TableOfAuthorities.KeepEntryFormatting Property (Word)

 **True** if formatting from table of authorities entries is applied to the entries in the specified table of authorities. Read/write **Boolean** .


## Syntax

 _expression_ . **KeepEntryFormatting**

 _expression_ An expression that returns a **[TableOfAuthorities](tableofauthorities-object-word.md)** object.


## Remarks

Corresponds to the \f switch for a Table of Authorities (TOA) field.


## Example

This example removes the formatting from the entries in the first table of authorities of the active document (the \f switch is added to the TOA field).


```vb
If ActiveDocument.TablesOfAuthorities.Count >= 1 Then 
 ActiveDocument.TablesOfAuthorities(1) _ 
 .KeepEntryFormatting = False 
End If
```


## See also


#### Concepts


[TableOfAuthorities Object](tableofauthorities-object-word.md)

