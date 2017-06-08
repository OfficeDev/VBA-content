---
title: Phonetics.Length Property (Excel)
keywords: vbaxl10.chm658075
f1_keywords:
- vbaxl10.chm658075
ms.prod: excel
api_name:
- Excel.Phonetics.Length
ms.assetid: 62f4c46d-2dc3-d8dc-b699-ca74eff1f77f
ms.date: 06/08/2017
---


# Phonetics.Length Property (Excel)

Returns a  **Long** value that represents the number of characters of phonetic text from the position you've specified with the **[Start](phonetics-start-property-excel.md)** property.


## Syntax

 _expression_ . **Length**

 _expression_ A variable that represents a **Phonetics** object.


## Example

This example returns the length of the second phonetic text string in the active cell.


```vb
ActiveCell.FormulaR1C1 = "東京都渋谷区代々木" 
ActiveCell.Phonetics.Add Start:=1, Length:=3, Text:="トウキョウト" 
ActiveCell.Phonetics.Add Start:=4, Length:=3, Text:="シブヤク" 
MsgBox ActiveCell.Phonetics(2).Length
```


## See also


#### Concepts


[Phonetics Object](phonetics-object-excel.md)

