---
title: Paragraph.IndentCharWidth Method (Word)
keywords: vbawd10.chm156696896
f1_keywords:
- vbawd10.chm156696896
ms.prod: word
api_name:
- Word.Paragraph.IndentCharWidth
ms.assetid: dba8182e-eb09-64dd-42c8-1e7e0e3af777
ms.date: 06/08/2017
---


# Paragraph.IndentCharWidth Method (Word)

Indents a paragraphs by a specified number of characters.


## Syntax

 _expression_ . **IndentCharWidth**( **_Count_** )

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of characters by which the specified paragraphs are to be indented.|

## Remarks

This method is equivalent to clicking the  **Increase Indent** button on the **Formatting** toolbar.


## Example

This example indents the first paragraph of the active document by 10 characters.


```vb
With ActiveDocument.Paragraphs(1) 
 .IndentCharWidth 10 
End With
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

