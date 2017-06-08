---
title: ODBCConnection.RefreshOnFileOpen Property (Excel)
keywords: vbaxl10.chm796082
f1_keywords:
- vbaxl10.chm796082
ms.prod: excel
api_name:
- Excel.ODBCConnection.RefreshOnFileOpen
ms.assetid: aa41bdde-c3c0-70ea-f3bc-99e641a306ac
ms.date: 06/08/2017
---


# ODBCConnection.RefreshOnFileOpen Property (Excel)

 **True** if the connection is automatically updated each time the workbook is opened. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **RefreshOnFileOpen**

 _expression_ A variable that represents an **ODBCConnection** object.


## Remarks

The connections are not automatically refreshed when you open the workbook by using the  **[Open](workbooks-open-method-excel.md)** method in Visual Basic. Use the **[Refresh](odbcconnection-refresh-method-excel.md)** method to refresh the data after the workbook is open.


## See also


#### Concepts


[ODBCConnection Object](odbcconnection-object-excel.md)

