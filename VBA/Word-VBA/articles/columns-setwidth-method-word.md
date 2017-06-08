---
title: Columns.SetWidth Method (Word)
keywords: vbawd10.chm155910345
f1_keywords:
- vbawd10.chm155910345
ms.prod: word
api_name:
- Word.Columns.SetWidth
ms.assetid: 42b9c3ce-5743-5143-f8e6-80bcbc0e206d
ms.date: 06/08/2017
---


# Columns.SetWidth Method (Word)

Sets the width of columns in a table.


## Syntax

 _expression_ . **SetWidth**( **_ColumnWidth_** , **_RulerStyle_** )

 _expression_ Required. A variable that represents a **[Columns](columns-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ColumnWidth_|Required| **Single**|The width of the specified column or columns, in points.|
| _RulerStyle_|Required| **[WdRulerStyle](wdrulerstyle-enumeration-word.md)**|Controls the way Word adjusts cell widths.|

## Remarks

The  **WdRulerStyle** behavior described above applies to left-aligned tables. The **WdRulerStyle** behavior for center- and right-aligned tables can be unexpected; in these cases, the **SetWidth** method should be used with care.


## See also


#### Concepts


[Columns Collection Object](columns-object-word.md)

