---
title: OLEDBConnection.ADOConnection Property (Excel)
keywords: vbaxl10.chm794073
f1_keywords:
- vbaxl10.chm794073
ms.prod: excel
api_name:
- Excel.OLEDBConnection.ADOConnection
ms.assetid: 918dd5eb-a9af-7466-92df-cae4e34676be
ms.date: 06/08/2017
---


# OLEDBConnection.ADOConnection Property (Excel)

Returns an ADO connection object if the PivotTable cache is connected to an OLE DB data source. Read-only.


## Syntax

 _expression_ . **ADOConnection**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

The  **ADOConnection** property exposes the Microsoft Excel connection to the data provider, allowing the user to write code within the context of the same session that Excel is using.

The  **ADOConnection** property is available only for sessions where the data source is an OLE DB data source. When there is no ADO session, the query will result in a run-time error. The **ADOConnection** property can be used for any OLE DB-based cache with ADO. The ADO connection object can be used with ADO MD for finding information about OLAP cubes on which the cache is based.


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

