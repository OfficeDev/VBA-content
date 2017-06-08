---
title: Cell.Split Method (Publisher)
keywords: vbapb10.chm5111844
f1_keywords:
- vbapb10.chm5111844
ms.prod: publisher
api_name:
- Publisher.Cell.Split
ms.assetid: 99bc34df-c8dc-90e5-4262-dbe0a9c9b61d
ms.date: 06/08/2017
---


# Cell.Split Method (Publisher)

Splits a merged table cell back into its constituent cells. Returns a  **[CellRange](cellrange-object-publisher.md)** object representing the constituent cells.


## Syntax

 _expression_. **Split**

 _expression_A variable that represents a  **Cell** object.


### Return Value

CellRange


## Remarks

If the specified cell is not a merged cell resulting from using the  **[Merge](cell-merge-method-publisher.md)** method, an error occurs.


## Example

The following example splits the first cell in the table in shape one on page one of the active publication into its constituent cells. Shape one must contain a table, the first cell of which is a merged cell, in order for this example to work.


```vb
Dim cllMerged As Cell 
 
Set cllMerged = ActiveDocument.Pages(1).Shapes(1).Table.Cells.Item(1) 
 
cllMerged.Split
```


