---
title: Range.ReadabilityStatistics Property (Word)
keywords: vbawd10.chm157155642
f1_keywords:
- vbawd10.chm157155642
ms.prod: word
api_name:
- Word.Range.ReadabilityStatistics
ms.assetid: c0dcf3e8-2c1a-3d23-48e9-4dfcd0d75893
ms.date: 06/08/2017
---


# Range.ReadabilityStatistics Property (Word)

Returns a  **ReadabilityStatistics** collection that represents the readability statistics for the specified document or range. Read-only.


## Syntax

 _expression_ . **ReadabilityStatistics**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning a Single Object from a Collection](http://msdn.microsoft.com/library/8c0b84c0-582b-32f7-68e0-6383d0661e74%28Office.15%29.aspx).


## Example

This example displays each readability statistic, along with its value, for document one.


```vb
For Each rs In Documents(1).ReadabilityStatistics 
 Msgbox rs.Name &; " - " &; rs.Value 
Next rs
```


## See also


#### Concepts


[Range Object](range-object-word.md)

