---
title: Cells.Delete Method (Word)
keywords: vbawd10.chm155844808
f1_keywords:
- vbawd10.chm155844808
ms.prod: word
api_name:
- Word.Cells.Delete
ms.assetid: 891c21b7-ef8d-9ba1-9408-6560dac146c7
ms.date: 06/08/2017
---


# Cells.Delete Method (Word)

Deletes a table cell or cells and optionally controls how the remaining cells are shifted.


## Syntax

 _expression_ . **Delete**( **_ShiftCells_** )

 _expression_ Required. A variable that represents a **[Cells](cells-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShiftCells_|Optional| **Variant**|The direction in which the remaining cells are to be shifted. Can be any  **[WdDeleteCells](wddeletecells-enumeration-word.md)** constant. If omitted, cells to the right of the last deleted cell are shifted left.|

## See also


#### Concepts


[Cells Collection Object](cells-object-word.md)

