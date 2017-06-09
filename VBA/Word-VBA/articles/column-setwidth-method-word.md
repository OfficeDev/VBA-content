---
title: Column.SetWidth Method (Word)
keywords: vbawd10.chm156172489
f1_keywords:
- vbawd10.chm156172489
ms.prod: word
api_name:
- Word.Column.SetWidth
ms.assetid: fd42d86d-53a4-c05d-81c3-add15cf05766
ms.date: 06/08/2017
---


# Column.SetWidth Method (Word)

Sets the width of a column in a table.


## Syntax

 _expression_ . **SetWidth**( **_ColumnWidth_** , **_RulerStyle_** )

 _expression_ Required. A variable that represents a **[Column](column-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ColumnWidth_|Required| **Single**|The width of the specified column or columns, in points.|
| _RulerStyle_|Required| **WdRulerStyle**|Controls the way Word adjusts cell widths.|

## Remarks

The  **[WdRulerStyle](wdrulerstyle-enumeration-word.md)** behavior described above applies to left-aligned tables. The **WdRulerStyle** behavior for center- and right-aligned tables can be unexpected; in these cases, the **SetWidth** method should be used with care.


## See also


#### Concepts


[Column Object](column-object-word.md)

