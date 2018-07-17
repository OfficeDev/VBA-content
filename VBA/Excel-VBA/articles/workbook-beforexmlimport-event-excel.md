---
title: Workbook.BeforeXmlImport Event (Excel)
keywords: vbaxl10.chm503097
f1_keywords:
- vbaxl10.chm503097
ms.prod: excel
api_name:
- Excel.Workbook.BeforeXmlImport
ms.assetid: a0a589c6-15f9-5599-c0b6-c6f881816ad6
ms.date: 06/08/2017
---


# Workbook.BeforeXmlImport Event (Excel)

Occurs before an existing XML data connection is refreshed or before new XML data is imported into a Microsoft Excel workbook.


## Syntax

 _expression_ . **BeforeXmlImport**( **_Map_** , **_Url_** , **_IsRefresh_** , **_Cancel_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Map_|Required| **[XmlMap](xmlmap-object-excel.md)**|The XML map that will be used to import data.|
| _Url_|Required| **String**|The location of the XML file to be imported.|
| _IsRefresh_|Required| **Boolean**| **True** if the event was triggered by refreshing an existing connection to XML data; **False** if the event was triggered by importing from a different data source.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the import or refresh operation.|

### Return Value

Nothing


