---
title: OLEDBConnection.IsConnected Property (Excel)
keywords: vbaxl10.chm794096
f1_keywords:
- vbaxl10.chm794096
ms.prod: excel
api_name:
- Excel.OLEDBConnection.IsConnected
ms.assetid: 3538c8bd-5027-8f48-d6b5-b18de0db4159
ms.date: 06/08/2017
---


# OLEDBConnection.IsConnected Property (Excel)

Returns  **True** if the **[MaintainConnection](oledbconnection-maintainconnection-property-excel.md)** property is ** True** . Returns **False** if it is not currently connected to its source. Read-only **Boolean** .


## Syntax

 _expression_ . **IsConnected**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

The  **IsConnected** property does not check to see if the connection is connected. Even if this property returns ** True** , sending commands to the provider could result in an error if the connection is no longer valid.


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

