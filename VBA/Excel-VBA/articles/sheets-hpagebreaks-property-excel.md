---
title: Sheets.HPageBreaks Property (Excel)
keywords: vbaxl10.chm152084
f1_keywords:
- vbaxl10.chm152084
ms.prod: excel
api_name:
- Excel.Sheets.HPageBreaks
ms.assetid: 5c7671c6-a00e-5183-db25-898509c7f8e8
ms.date: 06/08/2017
---


# Sheets.HPageBreaks Property (Excel)

Returns an  **[HPageBreaks](hpagebreaks-object-excel.md)** collection that represents the horizontal page breaks on the sheet. Read-only.


## Syntax

 _expression_ . **HPageBreaks**

 _expression_ A variable that represents a **Sheets** object.


## Remarks

There is a limit of 1026 horizontal page breaks per sheet.


## Example

This example displays the number of full-screen and print-area horizontal page breaks.


```vb
For Each pb in Worksheets(1).HPageBreaks 
 If pb.Extent = xlPageBreakFull Then 
 cFull = cFull + 1 
 Else 
 cPartial = cPartial + 1 
 End If 
Next 
MsgBox cFull &; " full-screen page breaks, " &; cPartial &; _ 
 " print-area page breaks"
```


## See also


#### Concepts


[Sheets Object](sheets-object-excel.md)

