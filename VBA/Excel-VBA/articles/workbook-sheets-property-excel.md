---
title: Workbook.Sheets Property (Excel)
keywords: vbaxl10.chm199152
f1_keywords:
- vbaxl10.chm199152
ms.prod: excel
api_name:
- Excel.Workbook.Sheets
ms.assetid: 45e4e19e-55ea-9615-231d-9435ba6d5a63
ms.date: 06/08/2017
---


# Workbook.Sheets Property (Excel)

Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the sheets in the specified workbook. Read-only **Sheets** object.


## Syntax

 _expression_ . **Sheets**

 _expression_ An expression that returns a **Workbook** object.


## Remarks

Using this property without an object qualifier is equivalent to using  `ActiveWorkbook.Sheets`.


## Example

This example creates a new worksheet and then places a list of the active workbook's sheet names in the first column.


```vb
Set newSheet = Sheets.Add(Type:=xlWorksheet) 
For i = 1 To Sheets.Count 
 newSheet.Cells(i, 1).Value = Sheets(i).Name 
Next i
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

