---
title: Range.SortAscending Method (Word)
keywords: vbawd10.chm157155497
f1_keywords:
- vbawd10.chm157155497
ms.prod: word
api_name:
- Word.Range.SortAscending
ms.assetid: 2e7cd40d-6ddd-c191-c082-1e5c852e80a7
ms.date: 06/08/2017
---


# Range.SortAscending Method (Word)

Sorts paragraphs or table rows in ascending alphanumeric order.


## Syntax

 _expression_ . **SortAscending**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

The first paragraph or table row is considered a header record and isn't included in the sort. Use the  **Sort** method to include the header record in a sort.This method offers a simplified form of sorting intended for mail merge data sources that contain columns of data. For most sorting tasks, use the **Sort** method.


## See also


#### Concepts


[Range Object](range-object-word.md)

