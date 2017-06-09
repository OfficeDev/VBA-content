---
title: FileExportConverter Object (Excel)
keywords: vbaxl10.chm862072
f1_keywords:
- vbaxl10.chm862072
ms.prod: excel
api_name:
- Excel.FileExportConverter
ms.assetid: 299f018e-0dfa-c101-7538-4a285918ac20
ms.date: 06/08/2017
---


# FileExportConverter Object (Excel)

Represents a file converter that is used to save files.


## Remarks

You cannot create a new file converter or add one to the  **[FileExportConverters](fileexportconverters-object-excel.md)** collection. **FileExportConverter** objects are added during installation of Microsoft Office or by installing supplemental file converters.


## Example

Use  **FileExportConverters** ( _Index_ ), where _Index_ is an integer, to return a single **FileExportConverter** object. The following example displays the extensions associated with the second Microsoft Excel worksheet converter in the collection.


```vb
MsgBox FileExportConverters(2).Extensions
```

The index number represents the position of the file converter in the  **FileExportConverters** collection. The following example displays the description for the first file converter in the collection.




```vb
MsgBox FileExportConvters(1).Description
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

