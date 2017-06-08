---
title: Application.SheetsInNewWorkbook Property (Excel)
keywords: vbaxl10.chm133207
f1_keywords:
- vbaxl10.chm133207
ms.prod: excel
api_name:
- Excel.Application.SheetsInNewWorkbook
ms.assetid: e2615d23-e0e0-34c4-0fd3-25f46a0d017b
ms.date: 06/08/2017
---


# Application.SheetsInNewWorkbook Property (Excel)

Returns or sets the number of sheets that Microsoft Excel automatically inserts into new workbooks. Read/write  **Long** .


## Syntax

 _expression_ . **SheetsInNewWorkbook**

 _expression_ A variable that represents an **Application** object.


## Example

This example displays the number of sheets automatically inserted into new workbooks.


```vb
MsgBox "Microsoft Excel inserts " &; _ 
 Application.SheetsInNewWorkbook &; _ 
 " sheet(s) in each new workbook"
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

