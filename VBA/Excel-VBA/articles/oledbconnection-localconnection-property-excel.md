---
title: OLEDBConnection.LocalConnection Property (Excel)
keywords: vbaxl10.chm794080
f1_keywords:
- vbaxl10.chm794080
ms.prod: excel
api_name:
- Excel.OLEDBConnection.LocalConnection
ms.assetid: 9f9e8aab-3804-1a30-3db1-4e453583ff1e
ms.date: 06/08/2017
---


# OLEDBConnection.LocalConnection Property (Excel)

Returns or sets the connection string to an offline cube file. Read/write  **String** .


## Syntax

 _expression_ . **LocalConnection**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

For a non-OLAP data source, the value of the  **LocalConnection** property is an empty string, and the **[UseLocalConnection](oledbconnection-uselocalconnection-property-excel.md)** property is set to **False** .

Setting the  **LocalConnection** property does not immediately initiate the connection to the data source. You must first use the **[Refresh](oledbconnection-refresh-method-excel.md)** method to make the connection and retrieve the data.

The value of the  **LocalConnection** property is used if the **UseLocalConnection** property is set to **True** . If the **UseLocalConnection** property is set to **False** , the **[Connection](oledbconnection-connection-property-excel.md)** property specifies the connection string for query tables based on sources other than local cube files.


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

