---
title: Table.ApplyStyleHeadingRows Property (Word)
keywords: vbawd10.chm156303562
f1_keywords:
- vbawd10.chm156303562
ms.prod: word
api_name:
- Word.Table.ApplyStyleHeadingRows
ms.assetid: 1c7fb6d5-9010-fded-d882-388d1e631da2
ms.date: 06/08/2017
---


# Table.ApplyStyleHeadingRows Property (Word)

 **True** for Microsoft Word to apply heading-row formatting to the first row of the selected table. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyStyleHeadingRows**

 _expression_ An expression that returns a **[Table](table-object-word.md)** object.


## Remarks

The specified table style must contain heading-row formatting to apply this formatting to a table.


## Example

This example formats the second table in the active document with the table style "Table Style 1" and removes formatting for the first and last rows and the first and last columns. This example assumes that a table style named "Table Style 1" exists and that it contains heading-row formatting.


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

