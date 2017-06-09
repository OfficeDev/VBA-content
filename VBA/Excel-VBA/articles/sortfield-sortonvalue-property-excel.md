---
title: SortField.SortOnValue Property (Excel)
keywords: vbaxl10.chm843074
f1_keywords:
- vbaxl10.chm843074
ms.prod: excel
api_name:
- Excel.SortField.SortOnValue
ms.assetid: eeaaf959-71d2-99a3-7e66-61744ad4709e
ms.date: 06/08/2017
---


# SortField.SortOnValue Property (Excel)

Retun the value on which the sort is performed for the specified  **SortField** object. Read-only.


## Syntax

 _expression_ . **SortOnValue**

 _expression_ A variable that represents a **SortField** object.


## Example

This sample sorts the data in column B on sheet1 by font color in an ascending


```vb
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear 
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add(Range("B1:B25"), _ 
 xlSortOnFontColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(0, 0, 0) 
 
With ActiveWorkbook.Worksheets("Sheet1").Sort 
 .SetRange Range("A1:B25") 
 .Header = xlGuess 
 .MatchCase = False 
 .Orientation = xlTopToBottom 
 .SortMethod = xlPinYin 
 .Apply 
End With
```

Cell color




```vb
SortOn = xlSortOnCellColor 
SortOnValue.Color = RGB(255, 255, 0)
```

Font color




```vb
SortOn = xlSortOnFontColor 
SortOnValue.Color = RGB(255, 255, 0)
```

Icons




```vb
SortOn = xlSortOnIcon 
SortOnValue.Color = RGB(255, 255, 0) 
SortField.SetIcon ActiveWorkbook.IconSets(1).Item(3)
```


## See also


#### Concepts


[SortField Object](sortfield-object-excel.md)

