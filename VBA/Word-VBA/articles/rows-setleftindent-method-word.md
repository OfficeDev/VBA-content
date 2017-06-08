---
title: Rows.SetLeftIndent Method (Word)
keywords: vbawd10.chm155975882
f1_keywords:
- vbawd10.chm155975882
ms.prod: word
api_name:
- Word.Rows.SetLeftIndent
ms.assetid: 4ce8093a-dcb9-4d2c-e841-176818d991b8
ms.date: 06/08/2017
---


# Rows.SetLeftIndent Method (Word)

Sets the indentation for a row or rows in a table.


## Syntax

 _expression_ . **SetLeftIndent**( **_LeftIndent_** , **_RulerStyle_** )

 _expression_ Required. A variable that represents a **[Rows](rows-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LeftIndent_|Required| **Single**|The distance (in points) between the current left edge of the specified row or rows and the desired left edge.|
| _RulerStyle_|Required| **WdRulerStyle**|Controls the way Microsoft Word adjusts the table when the left indent is changed. The  **WdRulerStyle** behavior applies to left-aligned tables. The **WdRulerStyle** behavior for center- and right-aligned tables can be unexpected; in these cases, use the **SetLeftIndent** method with care.|

## Example

This example creates a table in a new document and indents all rows in the table 0.5 inch (36 points). When you change the left indent, the cell widths are adjusted to preserve the right edge of the table.


```vb
Dim docNew As Document 
Dim tableNew As Table 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=3, NumColumns:=3) 
 
tableNew.Rows.SetLeftIndent LeftIndent:=InchesToPoints(0.5), _ 
 RulerStyle:=wdAdjustSameWidth
```

This example indents the first row in table one in the active document 18 points, and it narrows the width of the first column to preserve the position of the right edge of the table.




```vb
If ActiveDocument.Tables.Count >= 1 Then 
 ActiveDocument.Tables(1).Rows.SetLeftIndent LeftIndent:=18, _ 
 RulerStyle:=wdAdjustFirstColumn 
End If
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

