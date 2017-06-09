---
title: FileExportConverter.FileFormat Property (Excel)
keywords: vbaxl10.chm863075
f1_keywords:
- vbaxl10.chm863075
ms.prod: excel
api_name:
- Excel.FileExportConverter.FileFormat
ms.assetid: cdf0a922-ae9e-76b1-c8e5-228298920373
ms.date: 06/08/2017
---


# FileExportConverter.FileFormat Property (Excel)

Returns an integer that identifies the file format associated with the specified  **[FileExportConverter](fileexportconverter-object-excel.md)** object. Read-only.


## Syntax

 _expression_ . **FileFormat**

 _expression_ A variable that represents a **FileExportConverter** object.


## Example

The following example displays the file format identifier for the first file converter in the  **[FileExportConverters](fileexportconverters-object-excel.md)** collection.


```vb
Dim fcTemp As FileExportConverter 
Set fcTemp = FileExportConverters(1) 
 
MsgBox "The file format identifier for the file converter is: " &; fcTemp.FileFormat
```

The following example shows how to use the file format identifier as a parameter in the  **[SaveAs](workbook-saveas-method-excel.md)** method of the **Workbook** object to save a file using the first file converter in the **[FileExportConverters](fileexportconverters-object-excel.md)** collection.




```vb
ActiveWorkbook.SaveAs _ 
 Filename:="C:\temp\myFile.xyz", _ 
 FileFormat:=Application.FileExportConverters(1).FileFormat, _ 
 CreateBackup:=False
```


## See also


#### Concepts


[FileExportConverter Object](fileexportconverter-object-excel.md)

