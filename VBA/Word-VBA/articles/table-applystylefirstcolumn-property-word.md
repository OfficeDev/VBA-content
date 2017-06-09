---
title: Table.ApplyStyleFirstColumn Property (Word)
keywords: vbawd10.chm156303564
f1_keywords:
- vbawd10.chm156303564
ms.prod: word
api_name:
- Word.Table.ApplyStyleFirstColumn
ms.assetid: 9802ff74-321d-a44c-2cac-9f17b91210d2
ms.date: 06/08/2017
---


# Table.ApplyStyleFirstColumn Property (Word)

 **True** for Microsoft Word to apply first-column formatting to the first column of the specified table. Read/write **Boolean** .


## Syntax

 _expression_ . **ApplyStyleFirstColumn**

 _expression_ An expression that returns a **[Table](table-object-word.md)** object.


## Remarks

The specified table style must contain first-column formatting to apply this formatting to a table.


## Example

This example formats the second table in the active document with the table style "Table Style 1" and removes formatting for the first and last rows and the first and last columns. This example assumes that a table style named "Table Style 1" exists and that it contains first column formatting.


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

