---
title: ParagraphFormat.IndentCharWidth Method (Word)
keywords: vbawd10.chm156434752
f1_keywords:
- vbawd10.chm156434752
ms.prod: word
api_name:
- Word.ParagraphFormat.IndentCharWidth
ms.assetid: 52e9b6b1-15b3-5e03-7259-21d847c1d59c
ms.date: 06/08/2017
---


# ParagraphFormat.IndentCharWidth Method (Word)

Indents one or more paragraphs by a specified number of characters.


## Syntax

 _expression_ . **IndentCharWidth**( **_Count_** )

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of characters by which the specified paragraphs are to be indented.|

## Remarks

Using this method is equivalent to clicking the  **Increase Indent** button on the **Formatting** toolbar.


## Example

This example indents the first paragraph of the active document by 10 characters.


```
Selection.ParagraphFormat.IndentCharWidth 10
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

