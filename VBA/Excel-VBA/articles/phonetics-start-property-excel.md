---
title: Phonetics.Start Property (Excel)
keywords: vbaxl10.chm658074
f1_keywords:
- vbaxl10.chm658074
ms.prod: excel
api_name:
- Excel.Phonetics.Start
ms.assetid: 987613b4-7f33-7004-6abf-fb52061cb722
ms.date: 06/08/2017
---


# Phonetics.Start Property (Excel)

Returns the position that represents the first character of a phonetic text string in the specified cell. Read-only  **Long** .


## Syntax

 _expression_ . **Start**

 _expression_ A variable that represents a **Phonetics** object.


## Example

This example returns the starting position of the second phonetic text string in the active cell.


```vb
ActiveCell.FormulaR1C1 = "東京都渋谷区代々木" 
ActiveCell.Phonetics.Add Start:=1, Length:=3, Text:="トウキョウト" 
ActiveCell.Phonetics.Add Start:=4, Length:=3, Text:="シブヤク" 
MsgBox ActiveCell.Phonetics(2).Start
```


## See also


#### Concepts


[Phonetics Object](phonetics-object-excel.md)

