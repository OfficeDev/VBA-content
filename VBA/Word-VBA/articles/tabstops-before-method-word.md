---
title: TabStops.Before Method (Word)
keywords: vbawd10.chm156565606
f1_keywords:
- vbawd10.chm156565606
ms.prod: word
api_name:
- Word.TabStops.Before
ms.assetid: 7a6ff83f-a1cc-1f60-6a29-08bc1f94ef7f
ms.date: 06/08/2017
---


# TabStops.Before Method (Word)

Returns the next  **TabStop** object to the left of Position.


## Syntax

 _expression_ . **Before**( **_Position_** )

 _expression_ Required. A variable that represents a **[TabStops](tabstops-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Position_|Required| **Single**|A location on the ruler, in points.|

## Example

This example changes the alignment of the first custom tab stop in the first paragraph in the active document that's less than 2 inches from the left margin.


```vb
Dim tsTemp As TabStop 
 
Set tsTemp = ActiveDocument.Paragraphs(1) _ 
 .TabStops.Before(InchesToPoints(2)) 
tsTemp.Alignment = wdAlignTabCenter
```


## See also


#### Concepts


[TabStops Collection Object](tabstops-object-word.md)

