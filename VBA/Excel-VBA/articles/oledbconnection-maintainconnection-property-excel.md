---
title: OLEDBConnection.MaintainConnection Property (Excel)
keywords: vbaxl10.chm794081
f1_keywords:
- vbaxl10.chm794081
ms.prod: excel
api_name:
- Excel.OLEDBConnection.MaintainConnection
ms.assetid: ce913d74-d86d-006c-4def-da04a8c630b6
ms.date: 06/08/2017
---


# OLEDBConnection.MaintainConnection Property (Excel)

Returns  **True** if the connection to the specified data source is maintained after the refresh operation and until the workbook is closed. Read/write **Boolean** .


## Syntax

 _expression_ . **MaintainConnection**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

The default value is  **True** . If you anticipate frequent queries to a server, setting this property to **True** might improve performance by reducing reconnection time. Setting this property to **False** causes an open connection to be closed.


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

