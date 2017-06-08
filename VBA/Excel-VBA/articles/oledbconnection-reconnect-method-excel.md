---
title: OLEDBConnection.Reconnect Method (Excel)
keywords: vbaxl10.chm794105
f1_keywords:
- vbaxl10.chm794105
ms.prod: excel
api_name:
- Excel.OLEDBConnection.Reconnect
ms.assetid: 94f862a0-a42e-bd80-3e1c-9adc52414bfe
ms.date: 06/08/2017
---


# OLEDBConnection.Reconnect Method (Excel)

Drops and then reconnects the specified connection.


## Syntax

 _expression_ . **Reconnect**

 _expression_ A variable that represents an **[OLEDBConnection](oledbconnection-object-excel.md)** object.


## Example

The following code example causes the specified connection to reconnect.


```vb
ThisWorkbook.Connections(1).Reconnect
```


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

