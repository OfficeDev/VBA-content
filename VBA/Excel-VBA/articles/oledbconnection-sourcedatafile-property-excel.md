---
title: OLEDBConnection.SourceDataFile Property (Excel)
keywords: vbaxl10.chm794092
f1_keywords:
- vbaxl10.chm794092
ms.prod: excel
api_name:
- Excel.OLEDBConnection.SourceDataFile
ms.assetid: ddadf399-3f93-bd20-22cf-5f9303704218
ms.date: 06/08/2017
---


# OLEDBConnection.SourceDataFile Property (Excel)

Returns or sets a  **String** indicating the source data file for an OLE DB connection. Read/write.


## Syntax

 _expression_ . **SourceDataFile**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

For file-based data sources (for example, Access) the  **SourceDataFile** property contains a fully qualified path to the source data file. It is null for server-based data sources (for example, SQL Server). The **SourceDataFile** property is set to null if the **[Connection](oledbconnection-connection-property-excel.md)** property is changed programmatically.


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

