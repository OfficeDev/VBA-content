---
title: Application.WorkbookAfterXmlExport Event (Excel)
keywords: vbaxl10.chm504101
f1_keywords:
- vbaxl10.chm504101
ms.prod: excel
api_name:
- Excel.Application.WorkbookAfterXmlExport
ms.assetid: 9d542c67-4244-d018-4db6-3584f0caec7c
ms.date: 06/08/2017
---


# Application.WorkbookAfterXmlExport Event (Excel)

Occurs after Microsoft Excel saves or exports XML data from the specified workbook.


## Syntax

 _expression_ . **WorkbookAfterXmlExport**( **_Wb_** , **_Map_** , **_Url_** , **_Result_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The target workbook.|
| _Map_|Required| **[XmlMap](xmlmap-object-excel.md)**|The XML map that was used to save or export data.|
| _Url_|Required| **String**|The location of the XML file that was exported.|
| _Result_|Required| **[XlXmlExportResult](xlxmlexportresult-enumeration-excel.md)**| Indicates the results of the save or export operation.|

### Return Value

Nothing


## Remarks



| **XlXmlExportResult** can be one of the following **XlXmlExportResult** constants|
| **xlXmlExportSuccess** . The XML data file was successfully exported.|
| **xlXmlExportValidationFailed** . The contents of the XML data file do not match the specified schema map.|
Use the  **[AfterXmlExport](workbook-afterxmlexport-event-excel.md)** event if you want to perform an operation after XML data has been exported from a particular workbook.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)
[Application Object](application-object-excel.md)

