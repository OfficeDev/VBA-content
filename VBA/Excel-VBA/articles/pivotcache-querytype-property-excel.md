---
title: PivotCache.QueryType Property (Excel)
keywords: vbaxl10.chm227089
f1_keywords:
- vbaxl10.chm227089
ms.prod: excel
api_name:
- Excel.PivotCache.QueryType
ms.assetid: 61346ed2-1ada-a105-1894-b22861047c4f
ms.date: 06/08/2017
---


# PivotCache.QueryType Property (Excel)

Indicates the type of query used by Microsoft Excel to populate the PivotTable cache. Read-only  **[XlQueryType](xlquerytype-enumeration-excel.md)** .


## Syntax

 _expression_ . **QueryType**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks



| **XlQueryType** can be one of these **XlQueryType** constants.|
| **xlTextImport** . Based on a text file, for query tables only|
| **xlOLEDBQuery** . Based on an OLE DB query, including OLAP data sources|
| **xlWebQuery** . Based on a Web page, for query tables only|
| **xlADORecordset** . Based on an ADO recordset query|
| **xlDAORecordSet** . Based on a DAO recordset query, for query tables only|
| **xlODBCQuery** . Based on an ODBC data source|
You specify the data source in the prefix for the  **[Connection](pivotcache-connection-property-excel.md)** property's value.


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

