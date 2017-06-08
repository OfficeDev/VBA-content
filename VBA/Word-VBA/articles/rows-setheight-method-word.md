---
title: Rows.SetHeight Method (Word)
keywords: vbawd10.chm155975883
f1_keywords:
- vbawd10.chm155975883
ms.prod: word
api_name:
- Word.Rows.SetHeight
ms.assetid: 6c6dc63d-c17c-ad39-4d7a-bb5b608e776e
ms.date: 06/08/2017
---


# Rows.SetHeight Method (Word)

Sets the height of table rows.


## Syntax

 _expression_ . **SetHeight**( **_RowHeight_** , **_HeightRule_** )

 _expression_ Required. A variable that represents a **[Rows](rows-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RowHeight_|Required| **Single**|The height of the row or rows, in points.|
| _HeightRule_|Required| **WdRowHeightRule**|The rule for determining the height of the specified rows.|

## Example

This example creates a table and then sets the row height to 0.5 inch (36 points) for all rows in the table.


```vb
Set newDoc = Documents.Add 
Set aTable = _ 
 newDoc.Tables.Add(Range:=Selection.Range, NumRows:=3, _ 
 NumColumns:=3) 
aTable.Rows.SetHeight RowHeight:=InchesToPoints(0.5), _ 
 HeightRule:=wdRowHeightExactly
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

