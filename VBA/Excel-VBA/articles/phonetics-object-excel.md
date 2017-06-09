---
title: Phonetics Object (Excel)
keywords: vbaxl10.chm657072
f1_keywords:
- vbaxl10.chm657072
ms.prod: excel
api_name:
- Excel.Phonetics
ms.assetid: 77c0c55c-a181-c68a-24ed-e6bcaf514663
ms.date: 06/08/2017
---


# Phonetics Object (Excel)

A collection of all the  **[Phonetic](phonetic-object-excel.md)** objects in the specified range.


## Remarks

Each  **Phonetic** object contains information about a specific phonetic text string.


## Example

Use the  **[Phonetics](range-phonetics-property-excel.md)** property to return the **Phonetics** collection. The following example makes all phonetic text in the range A1:C4 visible.


```vb
Range("A1:C4").Phonetics.Visible = True
```

Use  **Phonetics** ( _index_ ), where _index_ is the index number of the phonetic text, to return a single **Phonetic** object. The following example sets the first phonetic text string in the active cell to "フリガナ".




```vb
ActiveCell.Phonetics(1).Text = "フリガナ"
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


