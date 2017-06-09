---
title: Range.SortDescending Method (Word)
keywords: vbawd10.chm157155498
f1_keywords:
- vbawd10.chm157155498
ms.prod: word
api_name:
- Word.Range.SortDescending
ms.assetid: 018f7566-29cb-ad7f-87ae-55f041ab72a1
ms.date: 06/08/2017
---


# Range.SortDescending Method (Word)

Sorts paragraphs in descending alphanumeric order.


## Syntax

 _expression_ . **SortDescending**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

The first paragraph or table row is considered a header record and isn't included in the sort. Use the  **Sort** method to include the header record in a sort.This method offers a simplified form of sorting intended for mail-merge data sources that contain columns of data. For most sorting tasks, use the **Sort** method.


## See also


#### Concepts


[Range Object](range-object-word.md)

