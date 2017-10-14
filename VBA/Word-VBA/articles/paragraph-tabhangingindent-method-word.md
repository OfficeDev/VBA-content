---
title: Paragraph.TabHangingIndent Method (Word)
keywords: vbawd10.chm156696880
f1_keywords:
- vbawd10.chm156696880
ms.prod: word
api_name:
- Word.Paragraph.TabHangingIndent
ms.assetid: bb29f459-4e38-e31d-ee18-ba061e5c116e
ms.date: 06/08/2017
---


# Paragraph.TabHangingIndent Method (Word)

Sets a hanging indent to a specified number of tab stops. .


## Syntax

 _expression_ . **TabHangingIndent**( **_Count_** )

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).|

## Remarks

You can also use this method to remove tab stops from a hanging indent if the value of Count is a negative number.


## Example

This example sets a hanging indent to the second tab stop for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).TabHangingIndent(2)
```

This example moves the hanging indent back one tab stop for the first paragraph in the active document.




```vb
ActiveDocument.Paragraphs(1).TabHangingIndent(-1)
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

