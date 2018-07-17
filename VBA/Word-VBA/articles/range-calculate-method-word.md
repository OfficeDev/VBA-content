---
title: Range.Calculate Method (Word)
keywords: vbawd10.chm157155500
f1_keywords:
- vbawd10.chm157155500
ms.prod: word
api_name:
- Word.Range.Calculate
ms.assetid: 756d6143-bf92-7669-f686-be23246c3a29
ms.date: 06/08/2017
---


# Range.Calculate Method (Word)

Calculates a mathematical expression within a range or selection. Returns the result as a  **Single** .


## Syntax

 _expression_ . **Calculate**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example inserts a mathematical expression at the beginning of the active document, calculates the expression, and then appends the results to the range. The result is "1 + 1 = 2".


```vb
Set myRange = ActiveDocument.Range(0, 0) 
myRange.InsertBefore "1 + 1 " 
myRange.InsertAfter "= " &; myRange.Calculate
```


## See also


#### Concepts


[Range Object](range-object-word.md)

