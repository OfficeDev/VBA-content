---
title: Row.SetLeftIndent Method (Word)
keywords: vbawd10.chm156238026
f1_keywords:
- vbawd10.chm156238026
ms.prod: word
api_name:
- Word.Row.SetLeftIndent
ms.assetid: 44e8d024-5a7c-b4cb-1f14-341954fe66c8
ms.date: 06/08/2017
---


# Row.SetLeftIndent Method (Word)

Sets the indentation for a row in a table.


## Syntax

 _expression_ . **SetLeftIndent**( **_LeftIndent_** , **_RulerStyle_** )

 _expression_ Required. A variable that represents a **[Row](row-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LeftIndent_|Required| **Single**|The distance (in points) between the current left edge of the specified row or rows and the desired left edge.|
| _RulerStyle_|Required| **WdRulerStyle**|Controls the way Word adjusts the table when the left indent is changed.|

## Remarks

The  **WdRulerStyle** behavior described above applies to left-aligned tables. The **WdRulerStyle** behavior for center- and right-aligned tables can be unexpected; in these cases, use the **SetLeftIndent** method with care.


## Example

This example creates a table in a new document and indents the first row 0.5 inch (36 points). When you change the left indent, the cell widths are adjusted to preserve the right edge of the table.


```vb
Dim docNew As Document 
Dim tableNew As Table 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=3, NumColumns:=3) 
 
tableNew.Rows(1).SetLeftIndent LeftIndent:=InchesToPoints(0.5), _ 
 RulerStyle:=wdAdjustSameWidth
```


## See also


#### Concepts


[Row Object](row-object-word.md)

