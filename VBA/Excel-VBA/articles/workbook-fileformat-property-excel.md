---
title: Workbook.FileFormat Property (Excel)
keywords: vbaxl10.chm199100
f1_keywords:
- vbaxl10.chm199100
ms.prod: excel
api_name:
- Excel.Workbook.FileFormat
ms.assetid: ef722c3c-90ea-9810-b853-a3fff19d5c60
ms.date: 06/08/2017
---


# Workbook.FileFormat Property (Excel)

Returns the file format and/or type of the workbook. Read-only  **[XlFileFormat](xlfileformat-enumeration-excel.md)** .


## Syntax

 _expression_ . **FileFormat**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example saves the active workbook in Normal file format if its current file format is Excel 97/95.


```vb
If ActiveWorkbook.FileFormat = xlExcel9795 Then 
 ActiveWorkbook.SaveAs fileFormat:=xlExcel12 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

