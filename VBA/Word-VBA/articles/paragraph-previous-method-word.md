---
title: Paragraph.Previous Method (Word)
keywords: vbawd10.chm156696901
f1_keywords:
- vbawd10.chm156696901
ms.prod: word
api_name:
- Word.Paragraph.Previous
ms.assetid: 0ccc928e-26c3-d5e6-ea99-a3d9776fbdd1
ms.date: 06/08/2017
---


# Paragraph.Previous Method (Word)

Returns the previous paragraph as a  **Paragraph** object.


## Syntax

 _expression_ . **Previous**( **_Count_** )

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Count_|Optional| **Variant**|The number of paragraphs by which you want to move back. The default value is 1.|

### Return Value

Paragraph


## Example

This example selects the paragraph that precedes the selection in the active document.


```
Selection.Previous(Unit:=wdParagraph, Count:=1).Select
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

