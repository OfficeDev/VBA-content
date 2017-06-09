---
title: Paragraphs.IndentFirstLineCharWidth Method (Word)
keywords: vbawd10.chm156762434
f1_keywords:
- vbawd10.chm156762434
ms.prod: word
api_name:
- Word.Paragraphs.IndentFirstLineCharWidth
ms.assetid: d0fc2250-8e3a-8a35-7d15-2bd9cc3653db
ms.date: 06/08/2017
---


# Paragraphs.IndentFirstLineCharWidth Method (Word)

Indents the first line of one or more paragraphs by a specified number of characters.


## Syntax

 _expression_ . **IndentFirstLineCharWidth**( **_Count_** )

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of characters by which the first line of each specified paragraph is to be indented.|

## Example

This example indents the first line of all paragraphs in the active document by 10 characters.


```vb
With ActiveDocument.Paragraphs 
 .IndentFirstLineCharWidth 10 
End With
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

