---
title: Table.ApplyStyleLastColumn Property (Word)
keywords: vbawd10.chm156303565
f1_keywords:
- vbawd10.chm156303565
ms.prod: word
api_name:
- Word.Table.ApplyStyleLastColumn
ms.assetid: db47720e-0351-c48d-6ebe-a149f2b8c84f
ms.date: 06/08/2017
---


# Table.ApplyStyleLastColumn Property (Word)

 **True** for Microsoft Word to apply last-column formatting to the last column of the specified table. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyStyleLastColumn**

 _expression_ An expression that returns a **[Table](table-object-word.md)** object.


## Remarks

The specified table style must contain last-column formatting to apply this formatting to a table.


## Example

This example formats the second table in the active document with the table style "Table Style 1" and removes formatting for the first and last rows and the first and last columns. This example assumes that a table style named "Table Style 1" exists and that it contains last-column formatting.


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

