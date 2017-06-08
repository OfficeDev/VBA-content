---
title: Application.WorkbookBeforeXmlImport Event (Excel)
keywords: vbaxl10.chm504098
f1_keywords:
- vbaxl10.chm504098
ms.prod: excel
api_name:
- Excel.Application.WorkbookBeforeXmlImport
ms.assetid: 33c7f386-9eec-6ba4-519e-9480ab2f5a31
ms.date: 06/08/2017
---


# Application.WorkbookBeforeXmlImport Event (Excel)

Occurs before an existing XML data connection is refreshed, or new XML data is imported into any open Microsoft Excel workbook.


## Syntax

 _expression_ . **WorkbookBeforeXmlImport**( **_Wb_** , **_Map_** , **_Url_** , **_IsRefresh_** , **_Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The target workbook.|
| _Map_|Required| **[XmlMap](xmlmap-object-excel.md)**|The XML map that will be used to import data.|
| _Url_|Required| **String**|The location of the XML file to be imported.|
| _IsRefresh_|Required| **Boolean**| **True** if the event was triggered by refreshing an existing connection to XML data, **False** if a new mapping will be created.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the import or refresh operation.|

### Return Value

Nothing


## Remarks

Use the  **[BeforeXmlImport](workbook-beforexmlimport-event-excel.md)** event if you want to capture XML data that is being imported or refreshed to a particular workbook.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)
[Application Object](application-object-excel.md)

