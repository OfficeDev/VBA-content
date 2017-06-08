---
title: Application.WorkbookBeforeXmlExport Event (Excel)
keywords: vbaxl10.chm504100
f1_keywords:
- vbaxl10.chm504100
ms.prod: excel
api_name:
- Excel.Application.WorkbookBeforeXmlExport
ms.assetid: 2c228d28-2d42-40b0-ee36-214bc720d78a
ms.date: 06/08/2017
---


# Application.WorkbookBeforeXmlExport Event (Excel)

Occurs before Microsoft Excel saves or exports XML data from the specified workbook.


## Syntax

 _expression_ . **WorkbookBeforeXmlExport**( **_Wb_** , **_Map_** , **_Url_** , **_Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The target workbook.|
| _Map_|Required| **[XmlMap](xmlmap-object-excel.md)**|The XML map that will be used to save or export data.|
| _Url_|Required| **String**|The location of the XML file to be exported.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the save or export operation.|

### Return Value

Nothing


## Remarks

Use the  **[BeforeXmlExport](workbook-beforexmlimport-event-excel.md)** event if you want to capture XML data that is being exported or saved from a particular workbook.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)
[Application Object](application-object-excel.md)

