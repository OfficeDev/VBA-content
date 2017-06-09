---
title: ODBCConnection.Connection Property (Excel)
keywords: vbaxl10.chm796077
f1_keywords:
- vbaxl10.chm796077
ms.prod: excel
api_name:
- Excel.ODBCConnection.Connection
ms.assetid: 2fcd1043-b088-cfde-9853-4a20da20be26
ms.date: 06/08/2017
---


# ODBCConnection.Connection Property (Excel)

Returns or sets a string that contains ODBC settings that enable Microsoft Excel to connect to an ODBC data source. Read/write  **Variant** .


## Syntax

 _expression_ . **Connection**

 _expression_ A variable that represents an **ODBCConnection** object.


## Remarks

Setting the  **Connection** property does not immediately initiate the connection to the data source. You must use the **[Refresh](odbcconnection-refresh-method-excel.md)** method to make the connection and retrieve the data.


## See also


#### Concepts


[ODBCConnection Object](odbcconnection-object-excel.md)

