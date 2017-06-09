---
title: ODBCConnection.SourceDataFile Property (Excel)
keywords: vbaxl10.chm796089
f1_keywords:
- vbaxl10.chm796089
ms.prod: excel
api_name:
- Excel.ODBCConnection.SourceDataFile
ms.assetid: f32c0eeb-e8f5-1a9f-13fd-ead4ad96381f
ms.date: 06/08/2017
---


# ODBCConnection.SourceDataFile Property (Excel)

Returns or sets a  **String** indicating the source data file for an ODBC connection. Read/write.


## Syntax

 _expression_ . **SourceDataFile**

 _expression_ A variable that represents an **ODBCConnection** object.


## Remarks

For file-based data sources (for example, Access) the  **SourceDataFile** property contains a fully qualified path to the source data file. It is null for server-based data sources (for example, SQL Server). The **SourceDataFile** property is set to null if the **[Connection](odbcconnection-connection-property-excel.md)** property is changed programmatically.


## See also


#### Concepts


[ODBCConnection Object](odbcconnection-object-excel.md)

