---
title: Workbook.Excel4IntlMacroSheets Property (Excel)
keywords: vbaxl10.chm199169
f1_keywords:
- vbaxl10.chm199169
ms.prod: excel
api_name:
- Excel.Workbook.Excel4IntlMacroSheets
ms.assetid: 70a8c8d0-1169-7c3d-904e-5a32a4693f45
ms.date: 06/08/2017
---


# Workbook.Excel4IntlMacroSheets Property (Excel)

Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the Microsoft Excel 4.0 international macro sheets in the specified workbook. Read-only.


## Syntax

 _expression_ . **Excel4IntlMacroSheets**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example displays the number of Microsoft Excel 4.0 international macro sheets in the active workbook.


```vb
MsgBox "There are " &; _ 
 ActiveWorkbook.Excel4IntlMacroSheets.Count &; _ 
 " Microsoft Excel 4.0 international macro sheets" &; _ 
 " in this workbook."
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

