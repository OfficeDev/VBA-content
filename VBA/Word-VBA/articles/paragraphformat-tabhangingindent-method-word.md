---
title: ParagraphFormat.TabHangingIndent Method (Word)
keywords: vbawd10.chm156434736
f1_keywords:
- vbawd10.chm156434736
ms.prod: word
api_name:
- Word.ParagraphFormat.TabHangingIndent
ms.assetid: 918cec1a-cd94-b2d1-bdbb-99fcbb648947
ms.date: 06/08/2017
---


# ParagraphFormat.TabHangingIndent Method (Word)

Sets a hanging indent to a specified number of tab stops. .


## Syntax

 _expression_ . **TabHangingIndent**( **_Count_** )

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).|

## Remarks

You can also use this method to remove tab stops from a hanging indent if the value of Count is a negative number.


## Example

This example sets a hanging indent of the selected paragraphs to the second tab stop.


```
Selection.ParagraphFormat.TabHangingIndent(2)
```

This example moves the hanging indent of the selected paragraphs back one tab stop.




```
Selection.ParagraphFormat.TabHangingIndent(-1)
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

