---
title: ODBCConnection.AlwaysUseConnectionFile Property (Excel)
keywords: vbaxl10.chm796092
f1_keywords:
- vbaxl10.chm796092
ms.prod: excel
api_name:
- Excel.ODBCConnection.AlwaysUseConnectionFile
ms.assetid: 445c7371-0ac6-b6f3-1a78-a406922d106f
ms.date: 06/08/2017
---


# ODBCConnection.AlwaysUseConnectionFile Property (Excel)

 **True** if the connection file is always used to establish connection to the data source. Read/write **Boolean** .


## Syntax

 _expression_ . **AlwaysUseConnectionFile**

 _expression_ A variable that represents an **ODBCConnection** object.


## Remarks

When this property is  **True** , the connection file will be used to establish the connection to the data source. If the connection embedded within the workbook is different from the external connection file, the embedded connection will be ignored and the external connection file will be the only version considered.


## See also


#### Concepts


[ODBCConnection Object](odbcconnection-object-excel.md)

