---
title: Workbook.AfterXmlImport Event (Excel)
keywords: vbaxl10.chm503098
f1_keywords:
- vbaxl10.chm503098
ms.prod: excel
api_name:
- Excel.Workbook.AfterXmlImport
ms.assetid: b43adf53-6b67-6127-e69d-6ea05f68b7f6
ms.date: 06/08/2017
---


# Workbook.AfterXmlImport Event (Excel)

Occurs after an existing XML data connection is refreshed or after new XML data is imported into the specified Microsoft Excel workbook.


## Syntax

 _expression_ . **AfterXmlImport**( **_Map_** , **_IsRefresh_** , **_Result_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Map_|Required| **[XmlMap](xmlmap-object-excel.md)**|The XML map that will be used to import data.|
| _IsRefresh_|Required| **Boolean**| **True** if the event was triggered by refreshing an existing connection to XML data; **False** if the event was triggered by importing from a different data source.|
| _Result_|Required| **[XlXmlImportResult](xlxmlimportresult-enumeration-excel.md)**|Indicates the results of the refresh or import operation.|

### Return Value

Nothing


## Remarks





| **XlXmlImportResult** can be one of the following **XlXmlImportResult** constants:|
| **xlXmlImportElementsTruncated** . The contents of the specified XML data file have been truncated because the XML data file is too large for the worksheet.|
| **xlXmlImportSuccess** . The XML data file was successfully imported.|
| **xlXmlImportValidationFailed** . The contents of the XML data file do not match the specified schema map.|

