---
title: TabStops.After Method (Word)
keywords: vbawd10.chm156565607
f1_keywords:
- vbawd10.chm156565607
ms.prod: word
api_name:
- Word.TabStops.After
ms.assetid: 4c081809-dfd9-b379-0f7b-ec1ef39eacfc
ms.date: 06/08/2017
---


# TabStops.After Method (Word)

Returns the next  **TabStop** object to the right of Position.


## Syntax

 _expression_ . **After**( **_Position_** )

 _expression_ Required. A variable that represents a **[TabStops](tabstops-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Position_|Required| **Single**|A location on the ruler, in points.|

## Example

This example changes the alignment of the first custom tab stop in the first paragraph in the active document that's more than 1 inch from the left margin.


```vb
Dim tabTemp as TabStop 
 
Set tabTemp = ActiveDocument.Paragraphs(1).TabStops _ 
 .After(InchesToPoints(1)) 
 
tabTemp.Alignment = wdAlignTabCenter
```


## See also


#### Concepts


[TabStops Collection Object](tabstops-object-word.md)

