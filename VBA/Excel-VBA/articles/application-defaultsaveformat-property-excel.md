---
title: Application.DefaultSaveFormat Property (Excel)
keywords: vbaxl10.chm133217
f1_keywords:
- vbaxl10.chm133217
ms.prod: EXCEL
api_name:
- Excel.Application.DefaultSaveFormat
ms.assetid: bb5c50db-8ba3-f79a-4577-f293ebc52b50
---


# Application.DefaultSaveFormat Property (Excel)

Returns or sets the default format for saving files. For a list of valid constants, see the  **[FileFormat](workbook-fileformat-property-excel.md)** property. Read/write **Long** .


## Syntax

 _expression_ . **DefaultSaveFormat**

 _expression_ A variable that represents an **Application** object.


## Example

This example sets the default format for saving files.


```vb
Application.DefaultSaveFormat = xlExcel4Workbook
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

