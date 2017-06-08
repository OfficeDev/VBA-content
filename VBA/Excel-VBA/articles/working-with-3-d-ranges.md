---
title: Working with 3-D Ranges
keywords: vbaxl10.chm5206007
f1_keywords:
- vbaxl10.chm5206007
ms.prod: excel
ms.assetid: f80e1976-6d24-8539-8dbe-f0072bbac60f
ms.date: 06/08/2017
---


# Working with 3-D Ranges

If you are working with the same range on more than one sheet, use the  **Array** function to specify two or more sheets to select. The following example formats the border of a 3-D range of cells.


```vb
Sub FormatSheets() 
 Sheets(Array("Sheet2", "Sheet3", "Sheet5")).Select 
 Range("A1:H1").Select 
 Selection.Borders(xlBottom).LineStyle = xlDouble 
End Sub
```


The following example applies the  **[FillAcrossSheets](worksheets-fillacrosssheets-method-excel.md)** method to transfer the formats and any data from the range on Sheet2 to the corresponding ranges on all the worksheets in the active workbook.




```vb
Sub FillAll() 
 Worksheets("Sheet2").Range("A1:H1") _ 
 .Borders(xlBottom).LineStyle = xlDouble 
 Worksheets.FillAcrossSheets (Worksheets("Sheet2") _ 
 .Range("A1:H1")) 
End Sub
```


