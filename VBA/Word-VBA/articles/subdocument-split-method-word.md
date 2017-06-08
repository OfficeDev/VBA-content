---
title: Subdocument.Split Method (Word)
keywords: vbawd10.chm159973477
f1_keywords:
- vbawd10.chm159973477
ms.prod: word
api_name:
- Word.Subdocument.Split
ms.assetid: f4548dbc-3b96-b271-8e71-0d436a1c3ecc
ms.date: 06/08/2017
---


# Subdocument.Split Method (Word)

Divides an existing subdocument into two subdocuments at the same level in master document view or outline view.


## Syntax

 _expression_ . **Split**( **_Range_** )

 _expression_ Required. A variable that represents a **[Subdocument](subdocument-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range that, when the subdocument is split, becomes a separate subdocument.|

## Remarks

The division is at the beginning of the specified range. An error occurs if the document isn't in either master document or outline view or if the range isn't at the beginning of a paragraph in a subdocument.


## Example

This example splits the selection from an existing subdocument into a separate subdocument.


```
Selection.Range.Subdocuments(1).Split Range:=Selection.Range
```


## See also


#### Concepts


[Subdocument Object](subdocument-object-word.md)

