---
title: Paragraph.IndentFirstLineCharWidth Method (Word)
keywords: vbawd10.chm156696898
f1_keywords:
- vbawd10.chm156696898
ms.prod: word
api_name:
- Word.Paragraph.IndentFirstLineCharWidth
ms.assetid: 165d59f2-4d66-3128-0ab9-e4a4074d4b7d
ms.date: 06/08/2017
---


# Paragraph.IndentFirstLineCharWidth Method (Word)

Indents the first line of one or more paragraphs by a specified number of characters.


## Syntax

 _expression_ . **IndentFirstLineCharWidth**( **_Count_** )

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of characters by which the first line of each specified paragraph is to be indented.|

## Example

This example indents the first line of the first paragraph in the active document by 10 characters.


```vb
With ActiveDocument.Paragraphs(1) 
 .IndentFirstLineCharWidth 10 
End With
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

