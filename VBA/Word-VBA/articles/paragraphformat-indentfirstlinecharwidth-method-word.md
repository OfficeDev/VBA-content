---
title: ParagraphFormat.IndentFirstLineCharWidth Method (Word)
keywords: vbawd10.chm156434754
f1_keywords:
- vbawd10.chm156434754
ms.prod: word
api_name:
- Word.ParagraphFormat.IndentFirstLineCharWidth
ms.assetid: 9531e607-4287-d4a3-de85-315e806d9b51
ms.date: 06/08/2017
---


# ParagraphFormat.IndentFirstLineCharWidth Method (Word)

Indents the first line of one or more paragraphs by a specified number of characters.


## Syntax

 _expression_ . **IndentFirstLineCharWidth**( **_Count_** )

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of characters by which the first line of each specified paragraph is to be indented.|

## Example

This example indents the first line of the first paragraph in the active document by 10 characters.


```
Selection.ParagraphFormat.IndentFirstLineCharWidth 10
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

