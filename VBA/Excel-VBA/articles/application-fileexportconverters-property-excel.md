---
title: Application.FileExportConverters Property (Excel)
keywords: vbaxl10.chm133318
f1_keywords:
- vbaxl10.chm133318
ms.prod: excel
api_name:
- Excel.Application.FileExportConverters
ms.assetid: 1b7289ea-344f-cc3d-ec31-04d4196533ff
ms.date: 06/08/2017
---


# Application.FileExportConverters Property (Excel)

Returns a  **[FileExportConverters](fileexportconverters-object-excel.md)** collection that represents all the file converters for saving files available to Microsoft Excel. Read-only.


## Syntax

 _expression_ . **FileExportConverters**

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


## Remarks

For more information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/f8a36459-f9dd-9f4c-ef7a-b188173434d5%28Office.15%29.aspx).


## Example

The following example displays the description for the first file converter in the  **[FileExportConverters](fileexportconverters-object-excel.md)** collection.


```vb
Dim fcTemp As FileExportConverter 
Set fcTemp = FileExportConverter(1) 
 
MsgBox fcTemp.Description
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

