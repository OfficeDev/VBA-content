---
title: FileImport Method
keywords: vbagr10.chm5207362
f1_keywords:
- vbagr10.chm5207362
ms.prod: excel
api_name:
- Excel.FileImport
ms.assetid: 30aafa3b-231c-0c08-07a7-e7494888b082
ms.date: 06/08/2017
---


# FileImport Method

Imports a specified file or range, or an entire sheet of data.

 _expression_. **FileImport( _FileName, Password, ImportRange, WorksheetName, OverwriteCells_)**

 _expression_ Required. An expression that returns an **Application** object.

 **FileName** Required **String**. The file that contains the data to be imported.
 **_Password_** Optional **Variant**. The password for the file to be imported, if the file is password protected.
 **_ImportRange_** Optional **Variant**. The range of cells to be imported, if the file to be imported is a Microsoft Excel worksheet or workbook. If this argument is omitted, the complete contents of the worksheet are imported.
 **_WorksheetName_** Optional **Variant**. The name of the worksheet to be imported, if the file to be imported is a Microsoft Excel workbook.
 **_OverwriteCells_** Optional **Variant**.  **True** to specify that the user be notified before imported data overwrites existing data on the specified datasheet. The default value is **True**.

## Example

This example imports data from the range A2:D5 on the worksheet named "MySheet" in the Microsoft Excel workbook named "mynums.xls."


```vb
With myChart.Application 
 .FileImport FileName:="C:\mynums.xls", _ 
 ImportRange:="A2:D5", WorksheetName:="MySheet", _ 
 OverwriteCells:=False 
End With
```


