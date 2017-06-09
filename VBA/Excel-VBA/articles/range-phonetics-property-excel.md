---
title: Range.Phonetics Property (Excel)
keywords: vbaxl10.chm144229
f1_keywords:
- vbaxl10.chm144229
ms.prod: excel
api_name:
- Excel.Range.Phonetics
ms.assetid: fdc05b76-b574-63ec-045a-42fdcfae8a9e
ms.date: 06/08/2017
---


# Range.Phonetics Property (Excel)

Returns the  **[Phonetics](phonetics-object-excel.md)** collection of the range. Read only.


## Syntax

 _expression_ . **Phonetics**

 _expression_ A variable that represents a **Range** object.


## Example

This example displays all of the  **Phonetic** objects in the active cell.


```vb
Set objPhon = ActiveCell.Phonetics 
With objPhon 
 For Each objPhonItem in objPhon 
 MsgBox "Phonetic object: " &; .Text 
 Next 
End With
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

