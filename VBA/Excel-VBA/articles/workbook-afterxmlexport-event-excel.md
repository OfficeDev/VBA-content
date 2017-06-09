---
title: Workbook.AfterXmlExport Event (Excel)
keywords: vbaxl10.chm503100
f1_keywords:
- vbaxl10.chm503100
ms.prod: excel
api_name:
- Excel.Workbook.AfterXmlExport
ms.assetid: fe1e0a53-9f4e-ac88-58f7-fe420e57cabd
ms.date: 06/08/2017
---


# Workbook.AfterXmlExport Event (Excel)

Occurs after Microsoft Excel saves or exports XML data from the specified workbook. 


## Syntax

 _expression_ . **AfterXmlExport**( **_Map_** , **_Url_** , **_Result_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Map_|Required| **[XmlMap](xmlmap-object-excel.md)**|The schema map that was used to save or export data.|
| _Url_|Required| **String**|The location of the XML file that was exported.|
| _Result_|Required| **XlXmlExportResult**|Indicates the results of the save or export operation.|

### Return Value

Nothing


## Remarks





| **XlXmlExportResult** can be one of the following **XlXmlExportResult** constants:|
| **xlXmlExportSuccess** . The XML data file was successfully exported.|
| **xlXmlExportValidationFailed** . The contents of the XML data file do not match the specified schema map.|

