---
title: OLEDBConnection.MakeConnection Method (Excel)
keywords: vbaxl10.chm794082
f1_keywords:
- vbaxl10.chm794082
ms.prod: excel
api_name:
- Excel.OLEDBConnection.MakeConnection
ms.assetid: ff618eae-1593-aabc-dbcb-427291caf923
ms.date: 06/08/2017
---


# OLEDBConnection.MakeConnection Method (Excel)

Establishes a connection for the specified OLE DB connection.


## Syntax

 _expression_ . **MakeConnection**

 _expression_ A variable that represents an **OLEDBConnection** object.


### Return Value

Nothing


## Remarks

The  **MakeConnection** method can be used when a connection drops and the user wants to reestablish the connection.

Various objects and methods might return a run-time error if the connection is dropped. Use of this method assures a connection before executing other objects or methods.




 **Note**  Microsoft Excel might drop a connection temporarily in the course of a session (unknown to the VBA programmer), so this method proves useful.

This method will result in a run-time error if the  **[MaintainConnection](oledbconnection-maintainconnection-property-excel.md)** property of the specified OLE DB connection has been set to **False** .


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

