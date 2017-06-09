---
title: Rows.AllowBreakAcrossPages Property (Word)
keywords: vbawd10.chm155975683
f1_keywords:
- vbawd10.chm155975683
ms.prod: word
api_name:
- Word.Rows.AllowBreakAcrossPages
ms.assetid: 6e5c0794-1dcc-9e70-7968-12a11adf6e56
ms.date: 06/08/2017
---


# Rows.AllowBreakAcrossPages Property (Word)

 **True** if the text in a table row or rows are allowed to split across a page break. Read/write **Long** .


## Syntax

 _expression_ . **AllowBreakAcrossPages**

 _expression_ Required. A variable that represents a **[Rows](rows-object-word.md)** collection.


## Remarks

This property can be  **True** , **False** or **wdUndefined** (only some of the specified text is allowed to split).


## Example

This example creates a new document with a 5x5 table and prevents the third row of the table from being split during pagination.


```vb
Dim docNew As Document 
Dim tableNew As Table 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=5, NumColumns:=5) 
 
tableNew.Rows(3).AllowBreakAcrossPages = False
```

This example determines whether the rows in the current table can be split across pages. If the insertion point isn't in a table, a message box is displayed.




```vb
Dim lngAllowBreak as Long 
 
Selection.Collapse Direction:=wdCollapseStart 
If Selection.Tables.Count = 0 Then 
 MsgBox "The insertion point is not in a table." 
Else 
 lngAllowBreak = Selection.Rows.AllowBreakAcrossPages 
End If
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

