---
title: Application.Sheets Property (Excel)
keywords: vbaxl10.chm132108
f1_keywords:
- vbaxl10.chm132108
ms.prod: excel
api_name:
- Excel.Application.Sheets
ms.assetid: 729a512a-8faa-3a7e-758b-ff76e7200662
ms.date: 06/08/2017
---


# Application.Sheets Property (Excel)

Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the sheets in the active workbook. Read-only **Sheets** object.


## Syntax

 _expression_ . **Sheets**

 _expression_ An expression that returns a **Application** object.


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


[Application Object](application-object-excel.md)

