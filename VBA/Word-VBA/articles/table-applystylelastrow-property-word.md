---
title: Table.ApplyStyleLastRow Property (Word)
keywords: vbawd10.chm156303563
f1_keywords:
- vbawd10.chm156303563
ms.prod: word
api_name:
- Word.Table.ApplyStyleLastRow
ms.assetid: 007ac0c4-bec8-9c48-99e2-017567415193
ms.date: 06/08/2017
---


# Table.ApplyStyleLastRow Property (Word)

 **True** for Microsoft Word to apply last-row formatting to the last row of the specified table. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyStyleLastRow**

 _expression_ An expression that returns a **[Table](table-object-word.md)** object.


## Remarks

The specified table style must contain last-row formatting to apply this formatting to a table.


## Example

This example formats the second table in the active document with the table style "Table Style 1" and removes formatting for the first and last rows and the first and last columns. This example assumes that a table style named "Table Style 1" exists and that it contains last-row formatting.


```vb
Sub TableStyles() 
 With ActiveDocument.Tables(2) 
 .Style = "Table Style 1" 
 .ApplyStyleFirstColumn = False 
 .ApplyStyleHeadingRows = False 
 .ApplyStyleLastColumn = False 
 .ApplyStyleLastRow = False 
 End With 
End Sub
```


## See also


#### Concepts


[Table Object](table-object-word.md)

