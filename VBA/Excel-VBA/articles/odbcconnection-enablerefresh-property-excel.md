---
title: ODBCConnection.EnableRefresh Property (Excel)
keywords: vbaxl10.chm796078
f1_keywords:
- vbaxl10.chm796078
ms.prod: excel
api_name:
- Excel.ODBCConnection.EnableRefresh
ms.assetid: 7d10e758-e92c-90c6-2f12-60b7b5f531ea
ms.date: 06/08/2017
---


# ODBCConnection.EnableRefresh Property (Excel)

 **True** if the connection can be refreshed by the user. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **EnableRefresh**

 _expression_ A variable that represents an **ODBCConnection** object.


## Remarks

The  **[RefreshOnFileOpen](odbcconnection-refreshonfileopen-property-excel.md)** property is ignored if the **EnableRefresh** property is set to **False** . For OLAP data sources, setting this property to **False** disables updates.


## See also


#### Concepts


[ODBCConnection Object](odbcconnection-object-excel.md)

