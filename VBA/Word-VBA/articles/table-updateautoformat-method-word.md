---
title: Table.UpdateAutoFormat Method (Word)
keywords: vbawd10.chm156303375
f1_keywords:
- vbawd10.chm156303375
ms.prod: word
api_name:
- Word.Table.UpdateAutoFormat
ms.assetid: d33f3b59-f05c-d51e-5f43-17d56af6693f
ms.date: 06/08/2017
---


# Table.UpdateAutoFormat Method (Word)

Updates the table with the characteristics of a predefined table format.


## Syntax

 _expression_ . **UpdateAutoFormat**

 _expression_ Required. A variable that represents a **[Table](table-object-word.md)** object.


## Remarks

As an example of how this method functions, if you apply a table format with  **AutoFormat** and then insert rows and columns, the table may no longer match the predefined look. **UpdateAutoFormat** restores the format.


## Example

This example creates a table, applies a predefined format to it, adds a row, and then reapplies the predefined format.


```vb
Dim docNew As Document 
Dim tableNew As Table 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Selection.Range, 5, 5) 
 
With tableNew 
 .AutoFormat Format:=wdTableFormatColumns1 
 .Rows.Add BeforeRow:=tableNew.Rows(1) 
End With 
MsgBox "Click OK to reapply autoformatting." 
tableNew.UpdateAutoFormat
```

This example restores the predefined format to the table that contains the insertion point.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Tables(1).UpdateAutoFormat 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Table Object](table-object-word.md)

