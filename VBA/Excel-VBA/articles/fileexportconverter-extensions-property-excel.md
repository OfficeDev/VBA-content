---
title: FileExportConverter.Extensions Property (Excel)
keywords: vbaxl10.chm863073
f1_keywords:
- vbaxl10.chm863073
ms.prod: excel
api_name:
- Excel.FileExportConverter.Extensions
ms.assetid: 448fdc36-4f11-1dff-98c1-797339e04ddb
ms.date: 06/08/2017
---


# FileExportConverter.Extensions Property (Excel)

Returns the file name extensions associated with the specified  **[FileExportConverter](fileexportconverter-object-excel.md)** object. Read-only **String** .


## Syntax

 _expression_ . **Extensions**

 _expression_ A variable that represents a **FileExportConverter** object.


## Example

The following example displays the file extensions for the first file converter in the  **[FileExportConverters](fileexportconverters-object-excel.md)** collection.


```vb
Dim fcTemp As FileExportConverter 
Set fcTemp = FileExportConverters(1) 
 
MsgBox "The file name extensions for the file converter are: " &; fcTemp.Extensions
```


## See also


#### Concepts


[FileExportConverter Object](fileexportconverter-object-excel.md)

