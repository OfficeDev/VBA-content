---
title: Paragraph.TabIndent Method (Word)
keywords: vbawd10.chm156696882
f1_keywords:
- vbawd10.chm156696882
ms.prod: word
api_name:
- Word.Paragraph.TabIndent
ms.assetid: 71878527-31e3-8d0b-7d12-3ced2cc6b5ab
ms.date: 06/08/2017
---


# Paragraph.TabIndent Method (Word)

Sets the left indent for the specified paragraphs to a specified number of tab stops. .


## Syntax

 _expression_ . **TabIndent**( **_Count_** )

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).|

## Remarks

You can also use this method to remove the indent if the value of Count is a negative number.


## Example

This example indents the first paragraph in the active document to the second tab stop.


```vb
ActiveDocument.Paragraphs(1).TabIndent(2)
```

This example moves the indent of the first paragraph in the active document back one tab stop.




```vb
ActiveDocument.Paragraphs(1).TabIndent(-1)
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

