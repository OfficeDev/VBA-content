---
title: TablesOfAuthorities.NextCitation Method (Word)
keywords: vbawd10.chm152174695
f1_keywords:
- vbawd10.chm152174695
ms.prod: word
api_name:
- Word.TablesOfAuthorities.NextCitation
ms.assetid: c0bfde51-ce49-1570-9599-515b43875dec
ms.date: 06/08/2017
---


# TablesOfAuthorities.NextCitation Method (Word)

Finds and selects the next instance of the text specified by the ShortCitation parameter.


## Syntax

 _expression_ . **NextCitation**( **_ShortCitation_** )

 _expression_ Required. A variable that represents a **[TablesOfAuthorities](tablesofauthorities-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShortCitation_|Required| **String**|The text of the short citation.|

## Example

This example selects the next citation in the active document that begins with "in re".


```vb
ActiveDocument.TablesOfAuthorities.NextCitation _ 
 ShortCitation:="in re"
```


## See also


#### Concepts


[TablesOfAuthorities Collection Object](tablesofauthorities-object-word.md)

