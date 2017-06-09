---
title: Application.WorkbookAfterXmlImport Event (Excel)
keywords: vbaxl10.chm504099
f1_keywords:
- vbaxl10.chm504099
ms.prod: excel
api_name:
- Excel.Application.WorkbookAfterXmlImport
ms.assetid: a58cc327-3776-fe5b-68d4-406269f30379
ms.date: 06/08/2017
---


# Application.WorkbookAfterXmlImport Event (Excel)

Occurs after an existing XML data connection is refreshed, or new XML data is imported into any open Microsoft Excel workbook.


## Syntax

 _expression_ . **WorkbookAfterXmlImport**( **_Wb_** , **_Map_** , **_IsRefresh_** , **_Result_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The target workbook.|
| _Map_|Required| **[XmlMap](xmlmap-object-excel.md)**|The XML map that was used to import data.|
| _IsRefresh_|Required| **Boolean**| **True** if the event was triggered by refreshing an existing connection to XML data, **False** if a new mapping was created.|
| _Result_|Required| **[XlXmlImportResult](xlxmlimportresult-enumeration-excel.md)**|Indicates the results of the refresh or import operation.|

### Return Value

Nothing


## Remarks



| **XlXmlImportResult** can be one of the following **XlXmlImportResult** constants|
| **xlXmlImportElementsTruncated** . The contents of the specified XML data file have been truncated because the XML data file is too large for the worksheet.|
| **xlXmlImportSuccess** . The XML data file was successfully imported.|
| **xlXmlImportValidationFailed** . The contents of the XML data file do not match the specified schema map.|
Use the  **[AfterXmlImport](workbook-afterxmlimport-event-excel.md)** event if you want to perform an operation after XML data has been imported into a particular workbook.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)
[Application Object](application-object-excel.md)

